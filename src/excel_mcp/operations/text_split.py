"""Text to Columns: split the text of one column into the columns to its right."""

from collections.abc import Iterator

from openpyxl.worksheet.worksheet import Worksheet

from excel_mcp.errors import InvalidArgumentError, LimitExceededError
from excel_mcp.operations.cells import store_typed, writable_cell
from excel_mcp.refs import MAX_COLUMN, CellRange
from excel_mcp.text_budget import TextBudget
from excel_mcp.workspace import sheet_names

_NAMED_DELIMITERS = {"tab": "\t", "semicolon": ";", "comma": ",", "space": " "}


def text_to_columns(
    sheet: Worksheet,
    area: CellRange,
    delimiters: list[str] | None,
    fixed_widths_chars: list[int] | None,
    qualifier: str,
    merge_delimiters: bool,
    max_cells: int,
) -> int:
    """Split each text cell; pieces are typed as in Excel (numbers, dates, formulas, text).

    Returns the number of cells split. Cells that are not text are left alone, and the cells
    the pieces go into must be empty.
    """
    if area.cols != 1:
        raise InvalidArgumentError(f"Text to columns needs a single column, not {area}.")
    if bool(delimiters) == bool(fixed_widths_chars):
        raise InvalidArgumentError("Give either delimiters or fixed_widths_chars.")
    separators = {_NAMED_DELIMITERS.get(name, name) for name in delimiters or []}
    if any(len(separator) != 1 for separator in separators):
        raise InvalidArgumentError(
            "Each delimiter is 'tab', 'semicolon', 'comma', 'space' or a single character."
        )
    if fixed_widths_chars and not (
        len(fixed_widths_chars) < MAX_COLUMN and min(fixed_widths_chars) >= 1
    ):
        raise InvalidArgumentError(
            f"fixed_widths_chars needs at most {MAX_COLUMN - 1:,} widths, each at least 1."
        )
    split = _split_rows(
        sheet, area, separators, fixed_widths_chars, qualifier, merge_delimiters, max_cells
    )
    names = sheet_names(sheet)
    for row, pieces in split:
        if area.min_col + len(pieces) - 1 > MAX_COLUMN:
            raise InvalidArgumentError(f"Row {row} would split past the last column.")
        for offset in range(1, len(pieces)):
            neighbour = writable_cell(sheet, row, area.min_col + offset)
            if neighbour.value is not None:
                raise InvalidArgumentError(
                    f"{neighbour.coordinate} is not empty; Excel would ask before replacing it. "
                    "Clear or move it first."
                )
    for row, pieces in split:
        for offset, piece in enumerate(pieces):
            store_typed(writable_cell(sheet, row, area.min_col + offset), piece, names)
    return len(split)


def _split_rows(
    sheet: Worksheet,
    area: CellRange,
    separators: set[str],
    fixed_widths_chars: list[int] | None,
    qualifier: str,
    merge_delimiters: bool,
    max_cells: int,
) -> list[tuple[int, list[str]]]:
    """The pieces of each text cell. Both the text read and the pieces made are counted as they
    come, so a refusal costs no more than the limits allow, however much text the cells hold."""
    budget = TextBudget(
        max_cells, "Splitting would read more than {limit} characters of text. Use fewer cells."
    )
    split = []
    pieces_made = 0
    for row in range(area.min_row, area.max_row + 1):
        cell = sheet.cell(row, area.min_col)
        if not isinstance(text := cell.value, str) or cell.data_type != "s":
            continue
        budget.spend(len(text))
        pieces = []
        for piece in _pieces(text, separators, fixed_widths_chars, qualifier, merge_delimiters):
            pieces_made += 1
            if pieces_made > max_cells:
                raise LimitExceededError(f"Splitting would fill more than {max_cells:,} cells.")
            pieces.append(piece)
        split.append((row, pieces))
    return split


def _pieces(
    text: str,
    separators: set[str],
    fixed_widths_chars: list[int] | None,
    qualifier: str,
    merge_delimiters: bool,
) -> Iterator[str]:
    if fixed_widths_chars:
        yield from _fixed_pieces(text, fixed_widths_chars)
        return
    field = ""
    quoted = False
    at_start = True
    first = True
    index = 0
    while index < len(text):
        char = text[index]
        if quoted:
            if char == qualifier and text[index + 1 : index + 2] == qualifier:
                field += char
                index += 1
            elif char == qualifier:
                quoted = False
            else:
                field += char
        elif at_start and qualifier and char == qualifier:
            quoted = True
            at_start = False
        elif char in separators:
            if field or first or not merge_delimiters:
                yield field
            first = False
            field, at_start = "", True
        else:
            field += char
            at_start = False
        index += 1
    if field or first or not merge_delimiters:
        yield field


def _fixed_pieces(text: str, widths: list[int]) -> Iterator[str]:
    start = 0
    for width in widths:
        if start >= len(text):
            return
        yield text[start : start + width]
        start += width
    if start < len(text):
        yield text[start:]
