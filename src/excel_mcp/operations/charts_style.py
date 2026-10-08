"""Applying chart options to openpyxl chart objects."""

from openpyxl.chart import BarChart, DoughnutChart, ScatterChart
from openpyxl.chart.data_source import NumFmt
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.marker import DataPoint, Marker
from openpyxl.chart.series import Series
from openpyxl.chart.shapes import GraphicalProperties

from excel_mcp.operations.charts_options import ROUND_TYPES, ChartOptions, ChartType

_LEGEND_POSITIONS = {"right": "r", "left": "l", "top": "t", "bottom": "b"}


def style_chart(chart, options: ChartOptions, chart_type: ChartType) -> None:
    chart.title = options.title
    chart.width = options.width_cm
    chart.height = options.height_cm
    if not options.show_legend:
        chart.legend = None
    elif options.legend_position:
        chart.legend.position = _LEGEND_POSITIONS[options.legend_position]
    if options.data_labels:
        for part in chart._charts:
            part.dataLabels = DataLabelList(
                showVal=True,
                showSerName=False,
                showCatName=False,
                showLegendKey=False,
                showPercent=False,
            )
    if options.grouping in ("stacked", "percent_stacked"):
        chart.grouping = "stacked" if options.grouping == "stacked" else "percentStacked"
        if isinstance(chart, BarChart):
            chart.overlap = 100
    if chart_type not in ROUND_TYPES:
        _style_axes(chart, options)
    if isinstance(chart, DoughnutChart):
        chart.holeSize = 50
    if isinstance(chart, ScatterChart) and (
        options.markers is not None or options.smooth is not None
    ):
        chart.scatterStyle = ("smooth" if options.smooth else "line") + (
            "" if options.markers is False else "Marker"
        )


def _style_axes(chart, options: ChartOptions) -> None:
    chart.x_axis.title = options.x_axis_title
    chart.y_axis.title = options.y_axis_title
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
    """Apply colors, markers and smoothing; `series` is in data column order."""
    for index, item in enumerate(series):
        if options.markers is not None:
            item.marker = Marker(symbol="circle" if options.markers else "none")
        if options.smooth is not None:
            item.smooth = options.smooth
        if index < len(colors) and chart_type not in ROUND_TYPES:
            _color_series(item, colors[index], chart_type)
    if chart_type in ROUND_TYPES:
        series[0].dPt = [
            DataPoint(idx=index, spPr=GraphicalProperties(solidFill=color))
            for index, color in enumerate(colors)
        ]


def _color_series(series: Series, color: str, chart_type: ChartType) -> None:
    if chart_type in ("line", "scatter", "radar"):
        series.graphicalProperties.line.solidFill = color
        if series.marker is not None and series.marker.symbol != "none":
            series.marker.graphicalProperties = GraphicalProperties(solidFill=color)
            series.marker.graphicalProperties.line.solidFill = color
    else:
        series.graphicalProperties.solidFill = color
