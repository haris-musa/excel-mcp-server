"""Building charts: series, plot groups (combos, secondary axes) and where the chart goes."""

# pyright: reportArgumentType=false, reportAttributeAccessIssue=false, reportOptionalMemberAccess=false
# openpyxl's stubs type its enum-like fields as Literals and its descriptors as plain attributes.

from openpyxl import Workbook
from openpyxl.chart import (
    AreaChart,
    BarChart,
    BubbleChart,
    DoughnutChart,
    LineChart,
    PieChart,
    RadarChart,
    ScatterChart,
)
from openpyxl.chart._chart import ChartBase
from openpyxl.chart.data_source import AxDataSource, NumDataSource, NumRef
from openpyxl.chart.series import Series, XYSeries
from openpyxl.worksheet.worksheet import Worksheet

from excel_mcp.errors import InvalidArgumentError
from excel_mcp.operations import chartex
from excel_mcp.operations import charts_axes as axes
from excel_mcp.operations.chart_index import (
    check_index,
    delete_chart,
    list_charts,
    replace_chart,
)
from excel_mcp.operations.charts_check import check_chart
from excel_mcp.operations.charts_data import Plot, SeriesIn, resolve_series
from excel_mcp.operations.charts_modern import check_modern, check_only_for, resolve_modern
from excel_mcp.operations.charts_options import (
    MODERN_TYPES,
    ROUND_TYPES,
    ChartOptions,
    ChartType,
    SeriesSpec,
    default_legend,
)
from excel_mcp.operations.charts_series import color_slices, style_series
from excel_mcp.operations.charts_style import style_chart
from excel_mcp.operations.formatting import parse_color
from excel_mcp.operations.sheets import validate_sheet_name
from excel_mcp.refs import cell_name, parse_cell
from excel_mcp.workspace import get_sheet

_CLASSES = {
    "column": BarChart,
    "bar": BarChart,
    "line": LineChart,
    "area": AreaChart,
    "pie": PieChart,
    "doughnut": DoughnutChart,
    "radar": RadarChart,
    "scatter": ScatterChart,
    "bubble": BubbleChart,
}
# The order Excel draws combo parts in: areas behind columns behind lines.
_DRAW_ORDER = list(_CLASSES)
_GROUPINGS = {"stacked": "stacked", "percent_stacked": "percentStacked"}
_XY = ("scatter", "bubble")


def create_chart(
    workbook: Workbook,
    sheet: str,
    anchor_cell: str | None,
    chart_type: ChartType,
    data_range: str | None,
    series_in: SeriesIn,
    series: list[SeriesSpec],
    categories: str | None,
    options: ChartOptions,
    index: int | None,
) -> str:
    """Add a chart (or replace chart `index`) and say where it went."""
    host = None if anchor_cell is None else get_sheet(workbook, sheet)
    anchor = None if anchor_cell is None else cell_name(*parse_cell(anchor_cell))
    if host is None:
        if index is not None:
            raise InvalidArgumentError(
                "index replaces a chart placed at anchor_cell; a chart sheet holds one chart, so "
                "delete the sheet and create it again."
            )
        validate_sheet_name(sheet, workbook.sheetnames)
    check_only_for(options, chart_type)
    if chart_type in MODERN_TYPES:
        if series_in == "rows":
            raise InvalidArgumentError(
                f"{chart_type} charts read their data in columns; list rows with `series`."
            )
        return _create_modern(
            workbook, host, anchor, chart_type, data_range, series, categories, options, index
        )
    plots = resolve_series(
        workbook,
        host.title if host else sheet,
        chart_type,
        data_range,
        series_in,
        series,
        categories,
    )
    check_chart(plots, chart_type, options)
    chart = build_chart(plots, chart_type, options)
    if host is None:
        workbook.create_chartsheet(sheet).add_chart(chart)
        return f"Added a {chart_type} chart on the new chart sheet {sheet!r}."
    if index is None:
        host.add_chart(chart, anchor)
        return f"Added a {chart_type} chart to {host.title} at {anchor}."
    moved = index > len(host._charts)
    replace_chart(host, index, chart, anchor)
    return _replaced(host, index, chart_type, anchor, len(host._charts) if moved else index)


def _create_modern(
    workbook: Workbook,
    host: Worksheet | None,
    anchor: str | None,
    chart_type: ChartType,
    data_range: str | None,
    series: list[SeriesSpec],
    categories: str | None,
    options: ChartOptions,
    index: int | None,
) -> str:
    if host is None or anchor is None:
        raise InvalidArgumentError(
            f"A {chart_type} chart sits on the sheet: give anchor_cell, e.g. 'E2'."
        )
    plots = resolve_modern(workbook, host.title, chart_type, data_range, series, categories)
    check_modern(plots, chart_type, options)
    if index is None:
        chartex.add(workbook, host, anchor, chart_type, plots, options, None)
        return f"Added a {chart_type} chart to {host.title} at {anchor}."
    check_index(host, index)
    classic = len(host._charts)
    if index > classic:
        old = chartex.modern_charts(host)[index - classic - 1]
        chartex.add(workbook, host, anchor, chart_type, plots, options, old)
        return _replaced(host, index, chart_type, anchor, index)
    delete_chart(host, index)
    chartex.add(workbook, host, anchor, chart_type, plots, options, None)
    return _replaced(host, index, chart_type, anchor, len(list_charts(host)))


def _replaced(host: Worksheet, index: int, chart_type: str, anchor: str | None, now: int) -> str:
    moved = f" It is now chart {now}." if now != index else ""
    return f"Replaced chart {index} of {host.title} with a {chart_type} chart at {anchor}.{moved}"


def build_chart(plots: list[Plot], chart_type: ChartType, options: ChartOptions) -> ChartBase:
    colors = [_color(color) for color in options.colors]
    groups: dict[tuple[str, bool], ChartBase] = {}
    for position, plot in enumerate(plots):
        kind = plot.spec.type or chart_type
        key = (kind, plot.spec.secondary_axis)
        if key not in groups:
            groups[key] = _new_group(kind, chart_type, options)
        series = _series(plot, kind)
        color = _color(plot.spec.color) or (colors[position] if position < len(colors) else None)
        style_series(series, plot.spec, kind, index=position, color=color, options=options)
        groups[key].series.append(series)
    if chart_type in ROUND_TYPES:
        color_slices(groups[(chart_type, False)].series[0], plots[0].points, colors)
    ordered = sorted(groups, key=lambda key: (key[1], _DRAW_ORDER.index(key[0])))
    chart = groups[ordered[0]]
    for key in ordered[1:]:
        chart += groups[key]
    style_chart(chart, options, options.legend or default_legend(chart_type, len(plots)))
    if chart_type not in ROUND_TYPES:
        # Areas fill the plot edge to edge, unless columns or lines share the chart.
        cross = "midCat" if {key[0] for key in groups} <= {"area", *_XY} else "between"
        _style_axes(chart, chart_type, options, cross)
        for key in ordered:
            if key[1]:
                axes.secondary_axes(
                    groups[key],
                    options.secondary_y_axis,
                    scatter=key[0] == "scatter",
                    cross=cross,
                )
    return chart


def _color(color: str | None) -> str | None:
    return parse_color(color)[2:] if color else None


def _new_group(kind: str, chart_type: str, options: ChartOptions) -> ChartBase:
    """An empty plot group; `kind` is what it draws, `chart_type` the chart's main type."""
    group = _CLASSES[kind]()
    group.varyColors = kind in ROUND_TYPES
    if options.grouping != "standard" and kind == chart_type:
        group.grouping = _GROUPINGS[options.grouping]
    if isinstance(group, BarChart):
        group.type = "bar" if kind == "bar" else "col"
        if group.grouping != "clustered":
            group.overlap = 100
        else:
            group.overlap = -27
        group.gapWidth = 182 if kind == "bar" else 219
    elif isinstance(group, LineChart):
        group.marker = True
    elif isinstance(group, ScatterChart):
        smooth = options.scatter_style.startswith("smooth")
        group.scatterStyle = "smoothMarker" if smooth else "lineMarker"
    elif isinstance(group, BubbleChart):
        group.bubbleScale = 100
        group.showNegBubbles = False
    elif isinstance(group, DoughnutChart):
        group.holeSize = 50
    return group


def _series(plot: Plot, kind: str) -> Series:
    values = NumDataSource(numRef=NumRef(f=plot.values))
    labels = AxDataSource(numRef=NumRef(f=plot.categories)) if plot.categories else None
    if kind in _XY:
        series = XYSeries()
        series.yVal, series.xVal = values, labels
        if plot.sizes:
            series.zVal = NumDataSource(numRef=NumRef(f=plot.sizes))
    else:
        series = Series()
        series.val, series.cat = values, labels
    series.tx = plot.name
    return series


def _style_axes(chart: ChartBase, chart_type: str, options: ChartOptions, cross: str) -> None:
    xy = chart_type in _XY
    bar = chart_type == "bar"
    x_spec = options.x_axis
    if bar:
        # Excel draws the first category at the bottom; list rows top-down as in the sheet,
        # with the value axis kept below the bars. `reverse` gives Excel's own order.
        x_spec = x_spec.model_copy(update={"reverse": not x_spec.reverse})
    axes.style_axis(
        chart.x_axis, x_spec, vertical=bar, default_gridlines=xy,
        line=(25000, 75000) if xy else (15000, 85000),
    )  # fmt: skip
    axes.style_axis(
        chart.y_axis, options.y_axis, vertical=not bar, default_gridlines=True,
        line=(25000, 75000) if xy else None,
    )  # fmt: skip
    chart.x_axis.axPos, chart.y_axis.axPos = ("l", "b") if bar else ("b", "l")
    chart.y_axis.crossBetween = cross
    if xy:
        chart.x_axis.crossBetween = cross
    if bar:
        chart.y_axis.crosses = "max" if x_spec.reverse else "autoZero"
