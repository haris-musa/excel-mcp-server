"""Charts built from a block of data on a sheet."""

from typing import Literal

from openpyxl.chart import AreaChart, BarChart, LineChart, PieChart, Reference, ScatterChart
from openpyxl.chart.series_factory import SeriesFactory
from openpyxl.worksheet.worksheet import Worksheet
from pydantic import BaseModel, Field

from excel_mcp.errors import InvalidArgumentError
from excel_mcp.refs import CellRange, cell_name, parse_cell, parse_range

ChartType = Literal["column", "bar", "line", "area", "pie", "scatter"]


class ChartOptions(BaseModel):
    """Chart titles, size and legend."""

    title: str | None = Field(default=None, description="Chart title.")
    x_axis_title: str | None = Field(default=None, description="Horizontal axis title.")
    y_axis_title: str | None = Field(default=None, description="Vertical axis title.")
    width_cm: float = Field(default=15, gt=0, le=100, description="Chart width in cm.")
    height_cm: float = Field(default=7.5, gt=0, le=100, description="Chart height in cm.")
    show_legend: bool = Field(default=True, description="Show the series legend.")


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

    if chart_type == "scatter":
        chart = _scatter(data_sheet, area)
    else:
        chart = _categorical(data_sheet, area, chart_type)
    chart.title = options.title
    chart.width = options.width_cm  # pyright: ignore[reportAttributeAccessIssue]
    chart.height = options.height_cm  # pyright: ignore[reportAttributeAccessIssue]
    if not options.show_legend:
        chart.legend = None
    if chart_type != "pie":
        chart.x_axis.title = options.x_axis_title
        chart.y_axis.title = options.y_axis_title
        # openpyxl marks axes as deleted by default, which hides them in current Excel.
        chart.x_axis.delete = False
        chart.y_axis.delete = False

    sheet.add_chart(chart, anchor)
    return str(area)


def _categorical(sheet: Worksheet, area: CellRange, chart_type: ChartType):
    chart = {
        "column": BarChart,
        "bar": BarChart,
        "line": LineChart,
        "area": AreaChart,
        "pie": PieChart,
    }[chart_type]()
    if chart_type == "bar":
        chart.type = "bar"
    series = Reference(
        sheet,
        min_col=area.min_col + 1,
        max_col=area.max_col,
        min_row=area.min_row,
        max_row=area.max_row,
    )
    labels = Reference(sheet, min_col=area.min_col, min_row=area.min_row + 1, max_row=area.max_row)
    chart.add_data(series, titles_from_data=True)
    chart.set_categories(labels)
    return chart


def _scatter(sheet: Worksheet, area: CellRange) -> ScatterChart:
    chart = ScatterChart()
    x_values = Reference(
        sheet, min_col=area.min_col, min_row=area.min_row + 1, max_row=area.max_row
    )
    for col in range(area.min_col + 1, area.max_col + 1):
        y_values = Reference(sheet, min_col=col, min_row=area.min_row, max_row=area.max_row)
        chart.series.append(SeriesFactory(y_values, x_values, title_from_data=True))
    return chart
