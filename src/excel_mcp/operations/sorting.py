"""Sorting the rows of a range the way Excel's Sort command does."""

import datetime as dt
import unicodedata
from copy import copy
from dataclasses import dataclass
from typing import Any, Literal, cast

from openpyxl.cell.cell import Cell, MergedCell
from openpyxl.comments import Comment
from openpyxl.formula.translate import Translator
from openpyxl.styles.cell_style import StyleArray
from openpyxl.utils.datetime import to_excel
from openpyxl.worksheet.hyperlink import Hyperlink
from openpyxl.worksheet.worksheet import Worksheet
from pydantic import Field

from excel_mcp.errors import InvalidArgumentError
from excel_mcp.formulas import storable_formula
from excel_mcp.inputs import InputModel
from excel_mcp.operations.columns import find_column
from excel_mcp.refs import CellRange, parse_range
from excel_mcp.workspace import sheet_names

Order = Literal["ascending", "descending"]

# Ascending order of value kinds; blanks always go last.
_NUMBER, _TEXT, _BOOLEAN, _ERROR, _BLANK = range(5)

SortValue = tuple[int, Any]


class SortKey(InputModel):
    column: str = Field(description="Header text (if has_header) or column letter, e.g. 'C'.")
    order: Order = "ascending"


@dataclass
class Saved:
    """A cell as it was before sorting."""

    value: Any  # whatever openpyxl stored: text, number, date, formula
    data_type: str
    style: StyleArray
    note: Comment | None
    link: Hyperlink | None
    origin: str


@dataclass
class _Row:
    cells: list[Saved]
    keys: list[SortValue]


def sort_range(
    sheet: Worksheet, ref: str, keys: list[SortKey], has_header: bool, max_cells: int
) -> int:
    """Sort a range's rows, moving values, styles, notes and links together.

    Returns the number of rows sorted.
    """
    area = parse_range(ref).within(max_cells)
    if any(
        area.overlaps(CellRange(m.min_row, m.min_col, m.max_row, m.max_col))
        for m in sheet.merged_cells.ranges
    ):
        raise InvalidArgumentError(
            f"{area} contains merged cells, which cannot be sorted. Unmerge them first."
        )
    first_row = area.min_row + has_header
    if first_row >= area.max_row:
        raise InvalidArgumentError(f"{area} has fewer than two rows to sort.")
    key_offsets = [find_column(sheet, area, key.column, has_header) - area.min_col for key in keys]

    rows = [
        _snapshot(cells, key_offsets)
        for cells in sheet.iter_rows(
            min_row=first_row, max_row=area.max_row, min_col=area.min_col, max_col=area.max_col
        )
    ]
    # Sorting is stable: sorting by the last key first lets earlier keys take precedence.
    for index in reversed(range(len(keys))):
        _sort_by_key(rows, index, descending=keys[index].order == "descending")
    write_rows(sheet, area.min_col, first_row, [row.cells for row in rows])
    return len(rows)


def _sort_by_key(rows: list[_Row], index: int, *, descending: bool) -> None:
    rows.sort(key=lambda row: row.keys[index], reverse=descending)
    rows.sort(key=lambda row: row.keys[index][0] == _BLANK)


def save_cells(row: tuple[Cell | MergedCell, ...]) -> list[Saved]:
    return [
        Saved(
            cell.value,
            cell.data_type,
            copy(cell._style),
            copy(cell.comment) if cell.comment else None,
            copy(cell.hyperlink) if cell.hyperlink else None,
            cell.coordinate,
        )
        for cell in row
    ]


def _snapshot(row: tuple[Cell | MergedCell, ...], key_offsets: list[int]) -> _Row:
    return _Row(save_cells(row), [_sort_value(row[offset]) for offset in key_offsets])


def _sort_value(cell: Cell | MergedCell) -> SortValue:
    if cell.data_type == "f":
        raise InvalidArgumentError(
            f"{cell.coordinate} is a formula, and its result is unknown until Excel "
            "recalculates. Sort by a column of values."
        )
    value = cell.value
    if value is None:
        return _BLANK, 0
    if cell.data_type == "e":
        return _ERROR, 0
    if isinstance(value, bool):
        return _BOOLEAN, value
    if isinstance(value, dt.datetime | dt.date | dt.time | dt.timedelta):
        return _NUMBER, float(to_excel(value))
    if isinstance(value, int | float):
        return _NUMBER, value
    return _TEXT, _text_key(str(value))


def _text_key(text: str) -> tuple[object, ...]:
    """Order text like Excel: ignoring case, hyphens and apostrophes, then by accents, then by
    how many hyphens and apostrophes there are. Symbols come before digits and letters."""
    decomposed = unicodedata.normalize("NFD", text.casefold())
    kept = "".join(c for c in decomposed if c not in "-'")
    base = "".join(c for c in kept if not unicodedata.combining(c))
    return (
        [(c.isalnum(), c) for c in base],
        kept,
        len(decomposed) - len(kept),
    )


def write_rows(sheet: Worksheet, first_col: int, first_row: int, rows: list[list[Saved]]) -> None:
    """Put saved rows back, relative references in formulas shifting as in a move."""
    names = sheet_names(sheet)
    for row_offset, row in enumerate(rows):
        for col_offset, saved in enumerate(row):
            # Merged cells are rejected before sorting, so this is always a real cell.
            cell = cast(Cell, sheet.cell(first_row + row_offset, first_col + col_offset))
            if saved.data_type == "f":
                if not isinstance(saved.value, str):
                    raise InvalidArgumentError(
                        f"{saved.origin} holds an array or data table formula, "
                        "which cannot be sorted."
                    )
                formula = Translator(saved.value, origin=saved.origin).translate_formula(
                    cell.coordinate
                )
                cell.value = storable_formula(formula, names)
            else:
                cell.value = saved.value
                cell.data_type = saved.data_type
            cell._style = copy(saved.style)
            cell.comment = saved.note
            cell.hyperlink = saved.link
