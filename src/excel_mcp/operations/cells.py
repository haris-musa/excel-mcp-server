"""Reading, writing, clearing, copying and searching cell contents."""

from collections.abc import Iterator
from copy import copy

from openpyxl.cell.cell import Cell, MergedCell
from openpyxl.formula.translate import Translator
from openpyxl.worksheet.formula import ArrayFormula
from openpyxl.worksheet.worksheet import Worksheet
from pydantic import BaseModel

from excel_mcp.errors import InvalidArgumentError, LimitExceededError
from excel_mcp.formulas import check_formula
from excel_mcp.refs import MAX_COLUMN, MAX_ROW, CellRange, cell_name, parse_cell, parse_range
from excel_mcp.values import CellValue, date_number_format, to_cell, to_json


class RangeData(BaseModel):
    sheet: str
    range: str
    values: list[list[CellValue]]
    truncated: bool
    next_range: str | None = None


class WriteResult(BaseModel):
    sheet: str
    range: str
    cells_written: int


class CellMatch(BaseModel):
    sheet: str
    cell: str
    value: CellValue


class FindResult(BaseModel):
    matches: list[CellMatch]
    truncated: bool


def stored_cells(sheet: Worksheet) -> Iterator[Cell]:
    """Cells that hold a value, row by row.

    Iterating the stored cells directly avoids creating empty cells for every
    position up to max_row x max_column, as iter_rows() would.
    """
    for _, cell in sorted(sheet._cells.items()):
        if cell.value is not None:
            yield cell


def used_range(sheet: Worksheet) -> CellRange:
    """The smallest range containing every cell that has a value.

    Worksheet.max_row/max_column also count cells that only carry formatting.
    """
    cells = list(stored_cells(sheet))
    if not cells:
        return CellRange(1, 1, 1, 1)
    rows = [cell.row for cell in cells]
    cols = [cell.column for cell in cells]
    return CellRange(min(rows), min(cols), max(rows), max(cols))


def writable_cell(sheet: Worksheet, row: int, col: int) -> Cell:
    if row > MAX_ROW or col > MAX_COLUMN:
        raise InvalidArgumentError(f"{cell_name(row, col)} is outside the worksheet limits.")
    cell = sheet.cell(row, col)
    if isinstance(cell, MergedCell):
        raise InvalidArgumentError(
            f"{cell.coordinate} is inside a merged range; write to its top-left cell."
        )
    return cell


def store_value(cell: Cell, value: CellValue) -> None:
    """Store a computed value, keeping text as text even if it starts with '='."""
    cell.value = value
    if isinstance(value, str):
        cell.data_type = "s"


def read_range(sheet: Worksheet, ref: str | None, max_cells: int) -> RangeData:
    target = parse_range(ref) if ref else used_range(sheet)
    if target.cols > max_cells:
        raise LimitExceededError(
            f"Range {target} has {target.cols} columns; at most {max_cells} cells can be "
            "read per call. Read fewer columns."
        )
    last_row = min(target.max_row, target.min_row + max_cells // target.cols - 1)
    rows = sheet.iter_rows(
        min_row=target.min_row,
        max_row=last_row,
        min_col=target.min_col,
        max_col=target.max_col,
        values_only=True,
    )
    read = CellRange(target.min_row, target.min_col, last_row, target.max_col)
    truncated = last_row < target.max_row
    remaining = CellRange(last_row + 1, target.min_col, target.max_row, target.max_col)
    return RangeData(
        sheet=sheet.title,
        range=str(read),
        values=[[to_json(value) for value in row] for row in rows],
        truncated=truncated,
        next_range=str(remaining) if truncated else None,
    )


def write_range(
    sheet: Worksheet, start_cell: str, rows: list[list[CellValue]], max_cells: int
) -> WriteResult:
    start_row, start_col = parse_cell(start_cell)
    if not any(rows):
        raise InvalidArgumentError("rows must contain at least one value.")
    count = sum(len(row) for row in rows)
    if count > max_cells:
        raise LimitExceededError(f"{count:,} cells exceeds the limit of {max_cells:,} per call.")
    width = max(len(row) for row in rows)
    written = CellRange(start_row, start_col, start_row + len(rows) - 1, start_col + width - 1)
    if written.max_row > MAX_ROW or written.max_col > MAX_COLUMN:
        raise InvalidArgumentError(f"Writing {written} would go past the worksheet limits.")

    converted = [[to_cell(value) for value in row] for row in rows]
    for row_offset, row in enumerate(converted):
        for col_offset, value in enumerate(row):
            cell = writable_cell(sheet, start_row + row_offset, start_col + col_offset)
            cell.value = value
            if number_format := date_number_format(value):
                cell.number_format = number_format
    return WriteResult(sheet=sheet.title, range=str(written), cells_written=count)


def clear_range(sheet: Worksheet, ref: str, contents: bool, formats: bool, max_cells: int) -> str:
    target = parse_range(ref).within(max_cells)
    for row in sheet.iter_rows(
        min_row=target.min_row,
        max_row=target.max_row,
        min_col=target.min_col,
        max_col=target.max_col,
    ):
        for cell in row:
            if isinstance(cell, MergedCell):
                continue
            if contents:
                cell.value = None
            if formats:
                cell.style = "Normal"
    return str(target)


def copy_range(
    source: Worksheet, ref: str, target: Worksheet, target_cell: str, max_cells: int
) -> str:
    """Copy values and styles; relative references in formulas shift the way Excel shifts them."""
    area = parse_range(ref).within(max_cells)
    target_row, target_col = parse_cell(target_cell)
    destination_area = CellRange(
        target_row, target_col, target_row + area.rows - 1, target_col + area.cols - 1
    )
    snapshot = [
        [(cell.value, cell.data_type, copy(cell._style), cell.coordinate) for cell in row]
        for row in source.iter_rows(
            min_row=area.min_row,
            max_row=area.max_row,
            min_col=area.min_col,
            max_col=area.max_col,
        )
    ]
    for row_offset, row in enumerate(snapshot):
        for col_offset, (value, data_type, style, origin) in enumerate(row):
            destination = writable_cell(target, target_row + row_offset, target_col + col_offset)
            if isinstance(value, ArrayFormula):
                raise InvalidArgumentError(
                    f"{origin} holds an array formula, which cannot be copied."
                )
            if data_type == "f":
                formula = Translator(str(value), origin=origin).translate_formula(
                    destination.coordinate
                )
                check_formula(formula)
                destination.value = formula
            else:
                destination.value = value
                destination.data_type = data_type
            destination._style = style
    return str(destination_area)


def find_cells(
    sheets: list[Worksheet],
    query: str,
    exact: bool,
    case_sensitive: bool,
    max_results: int,
) -> FindResult:
    def normalize(text: str) -> str:
        return text if case_sensitive else text.casefold()

    needle = normalize(query)
    matches: list[CellMatch] = []
    for sheet in sheets:
        for cell in stored_cells(sheet):
            text = normalize(str(cell.value))
            found = text == needle if exact else needle in text
            if not found:
                continue
            if len(matches) == max_results:
                return FindResult(matches=matches, truncated=True)
            matches.append(
                CellMatch(sheet=sheet.title, cell=cell.coordinate, value=to_json(cell.value))
            )
    return FindResult(matches=matches, truncated=False)
