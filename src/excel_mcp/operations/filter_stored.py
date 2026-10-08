"""Applying the criteria a file already stores, as Excel's Reapply does."""

from typing import cast

from openpyxl.worksheet import filters as xl
from openpyxl.worksheet.worksheet import Worksheet

from excel_mcp.errors import InvalidArgumentError
from excel_mcp.operations.comparison import FormulaResults, as_number, cell_value
from excel_mcp.operations.filter_criteria import (
    FilterColumn,
    FilterCondition,
    Operator,
    Test,
    average_filter,
    compare_filter,
    top_filter,
    values_filter,
)
from excel_mcp.refs import CellRange

_NAMES: dict[str, Operator] = {
    "equal": "equals",
    "notEqual": "not_equals",
    "greaterThan": "greater_than",
    "greaterThanOrEqual": "greater_or_equal",
    "lessThan": "less_than",
    "lessThanOrEqual": "less_or_equal",
}


def stored_test(
    sheet: Worksheet, area: CellRange, column: xl.FilterColumn, values: FormulaResults
) -> Test:
    """The test that shows the rows ``column``'s stored criteria show."""
    numbers = [
        n
        for row in range(area.min_row + 1, area.max_row + 1)
        if (n := as_number(cell_value(sheet.cell(row, area.min_col + column.colId), values)))
        is not None
    ]
    if column.filters is not None:
        return _values(column.filters)
    if column.customFilters is not None:
        conditions = [_condition(item) for item in column.customFilters.customFilter]
        combine = "and" if column.customFilters._and else "or"
        return compare_filter(conditions, combine)[1]
    if column.top10 is not None:
        top = column.top10.top is not False
        item = FilterColumn(
            column="",
            type="top" if top else "bottom",
            count=int(column.top10.val or 1),
            percent=bool(column.top10.percent),
        )
        return top_filter(item, numbers)[1]
    if column.dynamicFilter is not None and column.dynamicFilter.type in (
        "aboveAverage",
        "belowAverage",
    ):
        kind = "above_average" if column.dynamicFilter.type == "aboveAverage" else "below_average"
        return average_filter(kind, numbers)[1]
    if column.colorFilter is not None:
        return _color(sheet, column.colorFilter)
    raise InvalidArgumentError(
        f"The filter on column {column.colId + area.min_col} uses criteria that cannot be "
        "applied again here. Remove the filter first (set_sheet_layout auto_filter "
        "remove=true) and set it again afterwards."
    )


def _values(stored: xl.Filters) -> Test:
    wanted = [""] if stored.blank else []
    wanted += list(stored.filter)
    wanted += [f"{d.year:04}-{d.month:02}-{d.day:02}" for d in stored.dateGroupItem if d.day]
    return values_filter(wanted)[1]


def _condition(item: xl.CustomFilter) -> FilterCondition:
    name, text = _NAMES[item.operator or "equal"], str(item.val)
    if name in ("equals", "not_equals") and "*" in text.replace("~*", ""):
        negate = name == "not_equals"
        starts, ends = text.startswith("*"), text.endswith("*")
        body = text.strip("*").replace("~*", "*").replace("~?", "?").replace("~~", "~")
        kind = "contains" if starts and ends else "ends_with" if starts else "begins_with"
        return FilterCondition(
            operator=cast(Operator, f"not_{kind}" if negate else kind), value=body
        )
    text = text.replace("~*", "*").replace("~?", "?").replace("~~", "~")
    try:
        return FilterCondition(operator=name, value=float(text))
    except ValueError:
        return FilterCondition(operator=name, value=text)


def _color(sheet: Worksheet, stored: xl.ColorFilter) -> Test:
    styles = sheet.parent._differential_styles  # pyright: ignore[reportOptionalMemberAccess, reportAttributeAccessIssue]
    fill = styles[stored.dxfId].fill
    rgb = fill.fgColor.rgb if fill.fgColor.type == "rgb" else fill.bgColor.rgb
    return lambda cell, value: cell.fill.fill_type == "solid" and cell.fill.fgColor.rgb == rgb
