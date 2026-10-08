"""What a slicer does to a table: it filters a column, as the column's filter button does."""

from openpyxl.worksheet import filters as xl
from openpyxl.worksheet.table import Table, TableColumn
from openpyxl.worksheet.worksheet import Worksheet

from excel_mcp.errors import InvalidArgumentError
from excel_mcp.operations.comparison import FormulaResults, cell_value
from excel_mcp.operations.filter_criteria import FilterColumn, column_filter
from excel_mcp.operations.filters import table_area
from excel_mcp.refs import CellRange
from excel_mcp.text import quoted


def table_column(table: Table, name: str) -> tuple[int, TableColumn]:
    """The position (from 0) and definition of a column of the table."""
    for index, column in enumerate(table.tableColumns):
        if column.name.strip().casefold() == name.strip().casefold():
            return index, column
    raise InvalidArgumentError(
        f"Table {table.displayName!r} has no column {name!r}. "
        f"Columns: {quoted(c.name for c in table.tableColumns)}."
    )


def show_only(
    sheet: Worksheet, table: Table, column: str, items: list[str], values: FormulaResults
) -> None:
    """Filter the column to the items and hide the rows that fail."""
    area = table_area(table)
    index, _ = table_column(table, column)
    filters = table.autoFilter if table.autoFilter is not None else xl.AutoFilter(ref=str(area))
    if any(f.colId == index for f in filters.filterColumn):
        raise InvalidArgumentError(
            f"Column {column!r} of table {table.displayName!r} is already filtered. Remove "
            "that filter first (set_sheet_layout auto_filter)."
        )
    xml, _, test = column_filter(
        sheet, area, FilterColumn(column=column, type="values", values=items), values
    )
    filters.filterColumn.append(xml)
    table.autoFilter = filters
    _hide_failing_rows(sheet, area, index, test, values)


def _hide_failing_rows(
    sheet: Worksheet, area: CellRange, index: int, test, values: FormulaResults
) -> None:
    for row in range(area.min_row + 1, area.max_row + 1):
        cell = sheet.cell(row, area.min_col + index)
        if not test(cell, cell_value(cell, values)):
            sheet.row_dimensions[row].hidden = True


def shown_values(table: Table, index: int) -> list[str] | None:
    """The values a column's filter shows, or None when it has no values filter."""
    for column in table.autoFilter.filterColumn if table.autoFilter else []:
        if column.colId == index and column.filters is not None:
            return [*column.filters.filter, *([""] if column.filters.blank else [])]
    return None
