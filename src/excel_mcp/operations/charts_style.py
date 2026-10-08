"""Applying chart options to openpyxl chart objects."""

from openpyxl.chart import BarChart, DoughnutChart, ScatterChart
from openpyxl.chart.data_source import NumFmt
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.marker import DataPoint, Marker
from openpyxl.chart.series import Series
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.chart.title import Title, title_maker

from excel_mcp.operations.charts_options import ROUND_TYPES, ChartOptions, ChartType

_LEGEND_POSITIONS = {"right": "r", "left": "l", "top": "t", "bottom": "b"}
_GROUPINGS = {"stacked": "stacked", "percent_stacked": "percentStacked"}


def style_chart(chart, options: ChartOptions, chart_type: ChartType) -> None:
    chart.width = options.width_cm
    chart.height = options.height_cm
    chart.title = _title(options.title)
    if options.legend == "none":
        chart.legend = None
    else:
        # Without an explicit overlay flag, Excel draws the legend over the plot.
        chart.legend.overlay = False
        chart.legend.position = _LEGEND_POSITIONS[options.legend]
    if options.data_labels:
        for part in chart._charts:
            part.dataLabels = DataLabelList(
                showVal=True,
                showSerName=False,
                showCatName=False,
                showLegendKey=False,
                showPercent=False,
            )
    if options.grouping != "standard":
        chart.grouping = _GROUPINGS[options.grouping]
        if isinstance(chart, BarChart):
            chart.overlap = 100
    if chart_type not in ROUND_TYPES:
        _style_axes(chart, options)
    if isinstance(chart, DoughnutChart):
        chart.holeSize = 50
    if isinstance(chart, ScatterChart):
        chart.scatterStyle = "lineMarker"


def _title(text: str | None) -> Title | None:
    if text is None:
        return None
    title = title_maker(text)
    # Without an explicit overlay flag, Excel draws titles over the plot.
    title.overlay = False
    return title


def _style_axes(chart, options: ChartOptions) -> None:
    chart.x_axis.title = _title(options.x_axis_title)
    chart.y_axis.title = _title(options.y_axis_title)
    # openpyxl marks axes as deleted by default, which hides them in current Excel.
    chart.x_axis.delete = False
    chart.y_axis.delete = False
    chart.y_axis.scaling.min = options.y_axis_min
    chart.y_axis.scaling.max = options.y_axis_max
    if options.y_axis_number_format:
        chart.y_axis.numFmt = NumFmt(formatCode=options.y_axis_number_format, sourceLinked=False)


def style_series(
    series: list[Series], colors: list[str], chart_type: ChartType, options: ChartOptions
) -> None:
    """Apply markers, line shape and colors; `series` is in data column order."""
    for item in series:
        if chart_type == "line":
            item.marker = Marker(symbol="circle" if options.markers else "none")
            # Excel curves lines that carry no smooth flag.
            item.smooth = options.smooth
        elif chart_type == "scatter":
            # A scatter chart plots points, not a connecting line.
            item.marker = Marker(symbol="circle")
            item.graphicalProperties.line.noFill = True
            item.smooth = False
    if chart_type in ROUND_TYPES:
        series[0].dPt = [
            DataPoint(idx=index, spPr=GraphicalProperties(solidFill=color))
            for index, color in enumerate(colors)
        ]
        return
    for item, color in zip(series, colors, strict=False):
        _color_series(item, color, chart_type)


def _color_series(series: Series, color: str, chart_type: ChartType) -> None:
    if chart_type in ("line", "radar"):
        series.graphicalProperties.line.solidFill = color
    else:
        series.graphicalProperties.solidFill = color
    if series.marker is not None and series.marker.symbol != "none":
        series.marker.graphicalProperties = GraphicalProperties(solidFill=color)
        series.marker.graphicalProperties.line.solidFill = color
