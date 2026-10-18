"""Reading a CSV file into a worksheet."""

import csv
from dataclasses import dataclass
from pathlib import Path

from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
from openpyxl.worksheet.worksheet import Worksheet

from excel_mcp.errors import InvalidArgumentError, LimitExceededError
from excel_mcp.operations.cells import WriteResult, writable_cell
from excel_mcp.operations.csv_format import DELIMITER_CHARACTERS, FORMULA_STARTS, Encoding
from excel_mcp.operations.csv_values import DateOrder, typed_field
from excel_mcp.operations.date_width import FITTED_WIDTHS, widen_for_dates
from excel_mcp.refs import MAX_COLUMN, MAX_ROW, CellRange, parse_cell
from excel_mcp.values import store_value

MAX_FIELD_CHARACTERS = 32_767  # an Excel cell's limit
_SAMPLE_CHARACTERS = 64 * 1024


@dataclass(frozen=True)
class FieldRules:
    decimal: str
    date_order: DateOrder
    as_text: bool


def import_csv(
    sheet: Worksheet,
    file: Path,
    at: str,
    *,
    delimiter: str,
    encoding: Encoding,
    rules: FieldRules,
    max_cells: int,
) -> WriteResult:
    """Write the file's fields from ``at``, one row per record, streaming the file.

    ``delimiter`` is a name from `DELIMITER_CHARACTERS` or 'auto'.
    """
    first_row, first_col = parse_cell(at)
    cells = last_row = last_col = 0
    needed: dict[int, float] = {}
    codec = "utf-8-sig" if encoding == "utf-8" else encoding  # also reads UTF-8 without a BOM
    try:
        with file.open(encoding=codec, newline="") as stream:
            separator = _separator(stream.read(_SAMPLE_CHARACTERS), delimiter)
            stream.seek(0)
            for offset, record in enumerate(csv.reader(stream, delimiter=separator), start=1):
                row = first_row + offset - 1
                if record and (row > MAX_ROW or first_col + len(record) - 1 > MAX_COLUMN):
                    raise InvalidArgumentError(
                        f"Record {offset} does not fit on the worksheet from {at}."
                    )
                cells += len(record)
                if cells > max_cells:
                    raise LimitExceededError(
                        f"The file has more than {max_cells:,} cells, the limit per call. "
                        "Split the file or raise --max-cells."
                    )
                for index, text in enumerate(record):
                    column = first_col + index
                    if width := _store(sheet, row, column, text, offset, rules):
                        needed[column] = max(needed.get(column, 0), width)
                if any(record):
                    last_row, last_col = offset, max(last_col, len(record))
    except UnicodeError:
        raise InvalidArgumentError(
            f"The file is not valid {encoding}. Try another encoding, such as 'cp1252'."
        ) from None
    except csv.Error as error:
        raise InvalidArgumentError(f"The file is not valid CSV ({error}).") from None
    if not last_row:
        raise InvalidArgumentError("The file has no data.")
    widen_for_dates(sheet, needed)
    written = CellRange(first_row, first_col, first_row + last_row - 1, first_col + last_col - 1)
    return WriteResult(sheet=sheet.title, range=str(written), cells_written=cells)


def _separator(sample: str, delimiter: str) -> str:
    """The delimiter named, or for 'auto' the most frequent usual one in the first record."""
    if delimiter != "auto":
        return DELIMITER_CHARACTERS[delimiter]
    counts = dict.fromkeys(DELIMITER_CHARACTERS.values(), 0)
    quoted = False
    for character in sample:
        if character == '"':
            quoted = not quoted
        elif not quoted and character in "\r\n":
            break
        elif not quoted and character in counts:
            counts[character] += 1
    best = max(counts, key=lambda candidate: counts[candidate])
    return best if counts[best] else ","


def _store(
    sheet: Worksheet, row: int, column: int, text: str, record: int, rules: FieldRules
) -> float:
    """Store one field; returns the column width a date in it needs, else 0."""
    if not text:
        if (existing := sheet._cells.get((row, column))) is not None:  # pyright: ignore[reportAttributeAccessIssue]
            existing.value = None
        return 0
    if len(text) > MAX_FIELD_CHARACTERS or ILLEGAL_CHARACTERS_RE.search(text):
        raise InvalidArgumentError(
            f"A field in record {record} is longer than {MAX_FIELD_CHARACTERS:,} characters or "
            "holds control characters."
        )
    cell = writable_cell(sheet, row, column)
    formula_like = text.startswith(FORMULA_STARTS)
    text = text.replace("\r\n", "\n").replace("\r", "\n")  # Excel keeps line breaks as LF
    plain = cell.number_format == "General"  # openpyxl gives a date cell its own format
    value, number_format = (
        (text, None)
        if rules.as_text or cell.number_format == "@"
        else typed_field(text, rules.decimal, rules.date_order)
    )
    if isinstance(value, str):
        store_value(cell, value)
    else:
        cell.value = value
    # As Excel does for typed text, so that editing the cell later does not make a formula of it.
    cell.quotePrefix = isinstance(value, str) and formula_like
    if number_format and plain:
        cell.number_format = number_format
    return FITTED_WIDTHS.get(cell.number_format, 0)
