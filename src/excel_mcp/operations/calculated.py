"""Results for formula cells that Excel has not calculated yet."""

from openpyxl.styles.numbers import is_date_format
from openpyxl.utils.datetime import from_excel
from openpyxl.worksheet._read_only import ReadOnlyWorksheet

from excel_mcp.calc.engine import Engine
from excel_mcp.calc.values import ExcelError, Scalar, UncalculableError
from excel_mcp.operations import cells
from excel_mcp.operations.cells import RangeData
from excel_mcp.refs import CellRange, cell_name, parse_range
from excel_mcp.values import CellValue, to_json
from excel_mcp.workspace import Workspace, get_sheet, get_streamed_sheet

MAX_LISTED = 20


def read_calculated(
    workspace: Workspace, path: str, sheet: str, ref: str | None, max_cells: int
) -> RangeData:
    """Read values as Excel last stored them; calculate formulas that have no stored result.

    The window is streamed. Only when it holds such formulas is the workbook loaded in full,
    because evaluating a formula needs the cells it refers to, wherever they are.
    """
    with workspace.stream(path) as formulas:
        formula_sheet = get_streamed_sheet(formulas, sheet)
        # A used range found from stored results alone would miss formulas never calculated.
        window = str(cells.read_window(formula_sheet, ref))
    with workspace.stream(path, data_only=True) as stored:
        data = cells.read_range(get_streamed_sheet(stored, sheet), window, max_cells)
    area = parse_range(data.range)
    with workspace.stream(path) as formulas:
        formula_cells = _formula_cells(get_streamed_sheet(formulas, sheet), area)
    pending = [(row, col) for row, col in formula_cells if _stored(data, area, row, col) is None]
    if pending:
        _fill(workspace, path, sheet, data, pending)
    return data


def _formula_cells(sheet: ReadOnlyWorksheet, area: CellRange) -> list[tuple[int, int]]:
    found = []
    for row in sheet.iter_rows(
        min_row=area.min_row, max_row=area.max_row, min_col=area.min_col, max_col=area.max_col
    ):
        found.extend((cell.row, cell.column) for cell in row if cell.data_type == "f")
    return found


def _stored(data: RangeData, area: CellRange, row: int, col: int) -> CellValue:
    values = data.values
    line = values[row - area.min_row] if row - area.min_row < len(values) else []
    return line[col - area.min_col] if col - area.min_col < len(line) else None


def _fill(
    workspace: Workspace, path: str, sheet: str, data: RangeData, pending: list[tuple[int, int]]
) -> None:
    area = parse_range(data.range)
    unresolved: dict[str, str] = {}
    with workspace.read(path, data_only=True) as stored, workspace.read(path) as formulas:
        target = get_sheet(formulas, sheet)
        engine = Engine(stored, formulas)
        for row, col in pending:
            try:
                result = engine.calculate(target, row, col)
            except UncalculableError as reason:
                unresolved[cell_name(row, col)] = str(reason)
                continue
            number_format = target._cells[(row, col)].number_format
            _put(data, area, row, col, _json(result, number_format))
    if unresolved:
        listed = dict(list(unresolved.items())[:MAX_LISTED])
        if len(unresolved) > MAX_LISTED:
            listed["..."] = f"{len(unresolved) - MAX_LISTED} more"
        data.uncalculated = listed


def _put(data: RangeData, area: CellRange, row: int, col: int, value: CellValue) -> None:
    values = data.values
    while len(values) <= row - area.min_row:
        values.append([])
    line = values[row - area.min_row]
    while len(line) <= col - area.min_col:
        line.append(None)
    line[col - area.min_col] = value


def _json(result: Scalar, number_format: str) -> CellValue:
    if isinstance(result, ExcelError):
        return result.code
    if isinstance(result, bool | str):
        return result
    if is_date_format(number_format):
        return to_json(from_excel(result))
    number = float(f"{result:.15g}")
    return int(number) if number == int(number) and abs(number) < 1e15 else number
