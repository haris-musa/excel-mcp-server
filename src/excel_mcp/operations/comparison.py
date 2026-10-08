"""Cell values as filters and Remove Duplicates compare them."""

import datetime as dt

from openpyxl.cell.cell import Cell, MergedCell
from openpyxl.utils.datetime import to_excel

from excel_mcp.values import CellValue

FormulaResults = dict[str, CellValue | dt.datetime]


def cell_value(cell: Cell | MergedCell, results: FormulaResults) -> object:
    """The cell's value; for a formula, its calculated result."""
    return results.get(cell.coordinate) if cell.data_type == "f" else cell.value


def as_number(value: object) -> float | None:
    """The number Excel compares for a number or date (as a serial number), else None."""
    if isinstance(value, dt.date | dt.time | dt.timedelta):
        return float(to_excel(value))
    if isinstance(value, int | float) and not isinstance(value, bool):
        return float(value)
    return None
