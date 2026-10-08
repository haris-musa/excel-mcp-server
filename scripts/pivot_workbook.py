"""Builds the golden PivotTable workbook and reads back its cells (shared by generator and test)."""

import datetime as dt
from typing import Any

from openpyxl import Workbook
from openpyxl.utils.datetime import to_excel
from pivot_cases import SOURCE, SOURCE_RANGE

from excel_mcp.operations import pivot
from excel_mcp.operations.pivot import PivotRequest
from excel_mcp.operations.pivot_index import pivot_area, sheet_pivots
from excel_mcp.operations.pivot_options import CalculatedField, PivotField, PivotValue

MAX_CELLS = 1_000_000


def build_workbook(cases: list[dict[str, Any]]) -> Workbook:
    """One sheet per case, each with the PivotTable the server creates for it."""
    workbook = Workbook()
    source = workbook.worksheets[0]
    source.title = "Data"
    for row in SOURCE:
        source.append(row)
    for case in cases:
        request = PivotRequest(
            rows=case.get("rows", []),
            columns=case.get("columns", []),
            values=[PivotValue(**value) for value in case["values"]],
            filters=case.get("filters", []),
            fields=[PivotField(**field) for field in case.get("fields", [])],
            calculated=[CalculatedField(**field) for field in case.get("calculated_fields", [])],
            layout=case.get("layout", "tabular"),
            subtotals=case.get("subtotals", True),
            values_in=case.get("values_in", "columns"),
            name=case["name"],
        )
        target = workbook.create_sheet(case["name"])
        pivot.create_pivot(workbook, source, SOURCE_RANGE, target, "A1", request, MAX_CELLS)
    return workbook


def stored_grids(workbook: Workbook) -> dict[str, list[list[Any]]]:
    """The cells the server wrote for each PivotTable, as Excel's Value2 would show them."""
    grids = {}
    for sheet in workbook.worksheets[1:]:
        area = pivot_area(sheet_pivots(sheet)[0])
        grids[sheet.title] = [
            [_value(cell.value, cell.data_type) for cell in row]
            for row in sheet.iter_rows(
                min_row=area.min_row,
                max_row=area.max_row,
                min_col=area.min_col,
                max_col=area.max_col,
            )
        ]
    return grids


def _value(value: Any, data_type: str) -> Any:
    if data_type == "e":
        return value
    if isinstance(value, dt.datetime):
        return to_excel(value)
    return value


def same(expected: Any, actual: Any) -> bool:
    if expected in (None, "") and actual in (None, ""):
        return True
    if isinstance(expected, int | float) and isinstance(actual, int | float):
        return abs(expected - actual) <= 1e-9 * max(1.0, abs(expected), abs(actual))
    return expected == actual
