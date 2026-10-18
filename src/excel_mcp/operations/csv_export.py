"""Writing a worksheet range as CSV."""

import csv
import datetime as dt
import io
from typing import NamedTuple

from openpyxl.styles.numbers import is_date_format
from openpyxl.utils.datetime import to_excel

from excel_mcp.calc.textformat import format_number
from excel_mcp.calc.values import UncalculableError
from excel_mcp.errors import InvalidArgumentError, LimitExceededError
from excel_mcp.operations.calculated import read_calculated
from excel_mcp.operations.csv_format import DELIMITER_CHARACTERS, FORMULA_STARTS, Encoding
from excel_mcp.refs import CellRange, cell_name, parse_range
from excel_mcp.values import CellValue
from excel_mcp.workspace import Workspace, get_streamed_sheet

_EXCEL_PRECISION = 15


class CsvFile(NamedTuple):
    content: bytes
    range: str
    uncalculated: dict[str, str] | None


def export_csv(
    workspace: Workspace,
    path: str,
    sheet: str,
    ref: str | None,
    *,
    formatted: bool,
    delimiter: str,
    encoding: Encoding,
    escape_formulas: bool,
    max_cells: int,
) -> CsvFile:
    """Render a range's values (formula results included) as CSV with CRLF line ends."""
    data = read_calculated(workspace, path, sheet, ref, max_cells)
    if data.next_range:
        raise LimitExceededError(
            f"The range has more than {max_cells:,} cells, the limit per call. Export a "
            "smaller range or raise --max-cells."
        )
    area = parse_range(data.range)
    formats = _formats(workspace, path, sheet, area) if formatted else None
    buffer = io.StringIO(newline="")
    writer = csv.writer(buffer, delimiter=DELIMITER_CHARACTERS[delimiter], lineterminator="\r\n")
    for row in range(area.rows):
        line = data.values[row] if row < len(data.values) else []
        fields: list[str] = []
        for column in range(area.cols):
            value = line[column] if column < len(line) else None
            number_format = formats[row][column] if formats else None
            try:
                fields.append(_field(value, number_format, escape_formulas))
            except UncalculableError as reason:
                where = cell_name(area.min_row + row, area.min_col + column)
                raise InvalidArgumentError(
                    f"{where} cannot be written as displayed ({reason}). "
                    "Use formatted=false for its raw value."
                ) from None
        writer.writerow(fields)
    try:
        content = buffer.getvalue().encode(encoding)
    except UnicodeEncodeError:
        raise InvalidArgumentError(
            f"The data has characters that {encoding} cannot hold. Use 'utf-8'."
        ) from None
    return CsvFile(content, data.range, data.uncalculated)


def _formats(workspace: Workspace, path: str, sheet: str, area: CellRange) -> list[list[str]]:
    with workspace.stream(path) as workbook:
        rows = get_streamed_sheet(workbook, sheet).iter_rows(
            min_row=area.min_row, max_row=area.max_row, min_col=area.min_col, max_col=area.max_col
        )
        grid = [[cell.number_format or "General" for cell in row] for row in rows]
    # A sheet's stored rows can end early; the cells after them have no format of their own.
    grid += [["General"] * area.cols for _ in range(area.rows - len(grid))]
    return [row + ["General"] * (area.cols - len(row)) for row in grid]


def _field(value: CellValue, number_format: str | None, escape_formulas: bool) -> str:
    if value is None:
        return ""
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, str):
        shown = _shown_date(value, number_format) if number_format else None
        if shown is not None:
            return shown
        return "'" + value if escape_formulas and value.startswith(FORMULA_STARTS) else value
    if number_format:
        return format_number(float(value), number_format)
    return str(value) if isinstance(value, int) else f"{value:.{_EXCEL_PRECISION}g}"


def _shown_date(text: str, number_format: str) -> str | None:
    """A date or time read as ISO text, formatted; None if the cell is not one."""
    if not is_date_format(number_format):
        return None
    for parse in (dt.datetime.fromisoformat, dt.time.fromisoformat):
        try:
            return format_number(to_excel(parse(text)), number_format)
        except ValueError:
            continue
    return None
