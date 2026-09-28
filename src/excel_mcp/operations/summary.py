"""Grouped summary tables (a static alternative to pivot tables).

openpyxl cannot create real PivotTables, so this computes the aggregation in
Python and writes the result as ordinary cells.
"""

from collections import defaultdict
from statistics import fmean
from typing import Literal

from openpyxl.styles import Font
from openpyxl.worksheet.worksheet import Worksheet
from pydantic import BaseModel, Field

from excel_mcp.errors import InvalidArgumentError
from excel_mcp.operations.cells import store_value, writable_cell
from excel_mcp.refs import CellRange, parse_cell, parse_range
from excel_mcp.values import CellValue, to_json

Aggregation = Literal["sum", "average", "count", "min", "max"]


class SummaryValue(BaseModel):
    """A column to aggregate for each group."""

    field: str = Field(description="Header name of the column to aggregate.")
    aggregation: Aggregation = Field(default="sum", description="How to combine the values.")


def create_summary(
    source: Worksheet,
    source_ref: str,
    group_by: list[str],
    values: list[SummaryValue],
    target: Worksheet,
    target_cell: str,
    max_cells: int,
) -> str:
    if not group_by or not values:
        raise InvalidArgumentError("group_by and values each need at least one field.")
    headers, records = _read_records(source, parse_range(source_ref).within(max_cells))
    columns = {
        name: _column_index(headers, name) for name in [*group_by, *(v.field for v in values)]
    }

    groups: dict[tuple[CellValue, ...], list[list[CellValue]]] = defaultdict(list)
    for record in records:
        groups[tuple(record[columns[name]] for name in group_by)].append(record)

    output: list[list[CellValue]] = [[*group_by, *(f"{v.field} ({v.aggregation})" for v in values)]]
    for key in sorted(groups, key=lambda group: tuple(str(part) for part in group)):
        rows = groups[key]
        output.append(
            [
                *key,
                *(
                    _aggregate([row[columns[v.field]] for row in rows], v.aggregation)
                    for v in values
                ),
            ]
        )

    start_row, start_col = parse_cell(target_cell)
    for row_offset, row in enumerate(output):
        for col_offset, value in enumerate(row):
            cell = writable_cell(target, start_row + row_offset, start_col + col_offset)
            store_value(cell, value)
            if row_offset == 0:
                cell.font = Font(bold=True)
    return str(
        CellRange(start_row, start_col, start_row + len(output) - 1, start_col + len(output[0]) - 1)
    )


def _read_records(sheet: Worksheet, area: CellRange) -> tuple[list[str], list[list[CellValue]]]:
    rows = [
        [to_json(value) for value in row]
        for row in sheet.iter_rows(
            min_row=area.min_row,
            max_row=area.max_row,
            min_col=area.min_col,
            max_col=area.max_col,
            values_only=True,
        )
    ]
    if len(rows) < 2:
        raise InvalidArgumentError("Source data needs a header row and at least one data row.")
    headers = [str(header).strip() for header in rows[0]]
    return headers, [row for row in rows[1:] if any(value is not None for value in row)]


def _column_index(headers: list[str], name: str) -> int:
    folded = [header.casefold() for header in headers]
    if name.strip().casefold() not in folded:
        raise InvalidArgumentError(f"Field {name!r} not found. Available fields: {headers}.")
    return folded.index(name.strip().casefold())


def _aggregate(values: list[CellValue], aggregation: Aggregation) -> float | int | None:
    if aggregation == "count":
        return sum(value is not None for value in values)
    numbers = [
        value for value in values if isinstance(value, int | float) and not isinstance(value, bool)
    ]
    if not numbers:
        return None
    match aggregation:
        case "sum":
            return sum(numbers)
        case "average":
            return fmean(numbers)
        case "min":
            return min(numbers)
        case "max":
            return max(numbers)
