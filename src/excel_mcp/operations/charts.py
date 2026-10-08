"""Charts built from a block of data on a sheet."""

from openpyxl.chart import (
    AreaChart,
    BarChart,
    DoughnutChart,
    LineChart,
    PieChart,
    RadarChart,
    Reference,
    ScatterChart,
)
from openpyxl.chart.series import Series
from openpyxl.chart.series_factory import SeriesFactory
from openpyxl.worksheet.worksheet import Worksheet

from excel_mcp.errors import InvalidArgumentError
from excel_mcp.operations.charts_options import (
    ChartOptions,
    ChartType,
    check_options,
    parse_colors,
)
from excel_mcp.operations.charts_style import style_chart, style_series
from excel_mcp.refs import CellRange, cell_name, parse_cell, parse_range

_CATEGORICAL = {
    "column": BarChart,
    "bar": BarChart,
    "line": LineChart,
    "area": AreaChart,
    "pie": PieChart,
    "doughnut": DoughnutChart,
    "radar": RadarChart,
}


def create_chart(
    sheet: Worksheet,
    data_sheet: Worksheet,
    data_ref: str,
    chart_type: ChartType,
    anchor_cell: str,
    options: ChartOptions,
) -> str:
    area = parse_range(data_ref)
    if area.rows < 2 or area.cols < 2:
        raise InvalidArgumentError(
            "Chart data needs a header row and a label column plus at least one series, "
            "e.g. 'A1:C10' with labels in A and series in B and C."
        )
    anchor = cell_name(*parse_cell(anchor_cell))
    check_options(options, chart_type)
    slots = area.rows - 1 if chart_type in ("pie", "doughnut") else area.cols - 1
    colors = parse_colors(options, chart_type, slots)

    if chart_type == "scatter":
        chart, series = _scatter(data_sheet, area)
    else:
        chart, series = _categorical(data_sheet, area, chart_type, options)
    style_chart(chart, options, chart_type)
    style_series(series, colors, chart_type, options)

    sheet.add_chart(chart, anchor)
    return str(area)


def _categorical(sheet: Worksheet, area: CellRange, chart_type: ChartType, options: ChartOptions):
    chart = _CATEGORICAL[chart_type]()
    if chart_type == "bar":
        chart.type = "bar"
    secondary = _secondary_columns(sheet, area, options)
    line = LineChart() if secondary else None
    series: list[Series] = []
    for col in range(area.min_col + 1, area.max_col + 1):
        target = line if col in secondary and line else chart
        data = Reference(sheet, min_col=col, min_row=area.min_row, max_row=area.max_row)
        target.add_data(data, titles_from_data=True)
        series.append(target.series[-1])
    labels = Reference(sheet, min_col=area.min_col, min_row=area.min_row + 1, max_row=area.max_row)
    chart.set_categories(labels)
    if line:
        line.set_categories(labels)
        line.y_axis.axId = 200
        line.y_axis.crosses = "max"
        line.y_axis.delete = False
        line.y_axis.majorGridlines = None
        for item in line.series:
            item.smooth = False
        chart += line
    return chart, series


def _secondary_columns(sheet: Worksheet, area: CellRange, options: ChartOptions) -> set[int]:
    """Column numbers of the headers named in secondary_line_columns."""
    names = options.secondary_line_columns
    headers = {
        str(sheet.cell(area.min_row, col).value): col
        for col in range(area.min_col + 1, area.max_col + 1)
    }
    unknown = [name for name in names if name not in headers]
    if unknown:
        raise InvalidArgumentError(
            f"secondary_line_columns {unknown} not found among the series headers: "
            f"{', '.join(repr(header) for header in headers)}."
        )
    columns = {headers[name] for name in names}
    if len(columns) == len(headers):
        raise InvalidArgumentError(
            "secondary_line_columns cannot include every series; leave at least one as "
            "columns or bars."
        )
    return columns


def _scatter(sheet: Worksheet, area: CellRange):
    chart = ScatterChart()
    x_values = Reference(
        sheet, min_col=area.min_col, min_row=area.min_row + 1, max_row=area.max_row
    )
    for col in range(area.min_col + 1, area.max_col + 1):
        y_values = Reference(sheet, min_col=col, min_row=area.min_row, max_row=area.max_row)
        chart.series.append(SeriesFactory(y_values, x_values, title_from_data=True))
    return chart, list(chart.series)
