"""Series formatting: colors, lines, markers, labels, trendlines and error bars."""

# pyright: reportArgumentType=false, reportAttributeAccessIssue=false, reportOptionalMemberAccess=false
# openpyxl's stubs type its enum-like fields as Literals and its descriptors as plain attributes.

from openpyxl.chart.data_source import NumFmt
from openpyxl.chart.error_bar import ErrorBars as XmlErrorBars
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.marker import DataPoint, Marker
from openpyxl.chart.series import Series
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.chart.trendline import Trendline as XmlTrendline
from openpyxl.chart.trendline import TrendlineLabel
from openpyxl.drawing.colors import ColorChoice
from openpyxl.drawing.line import LineProperties

from excel_mcp.operations import charts_look as look
from excel_mcp.operations.charts_options import ChartOptions, SeriesSpec, Trendline

_EMU_PER_POINT = 12700
_BUBBLE_ALPHA = 75000
_LABEL_POSITIONS = {
    "center": "ctr",
    "inside_end": "inEnd",
    "inside_base": "inBase",
    "outside_end": "outEnd",
    "above": "t",
    "below": "b",
    "left": "l",
    "right": "r",
    "best_fit": "bestFit",
}
_TRENDLINES = {
    "linear": "linear",
    "exponential": "exp",
    "logarithmic": "log",
    "polynomial": "poly",
    "power": "power",
    "moving_average": "movingAvg",
}
_ERROR_VALUES = {
    "fixed": "fixedVal",
    "percent": "percentage",
    "std_dev": "stdDev",
    "std_error": "stdErr",
}


def style_series(
    series: Series,
    spec: SeriesSpec,
    kind: str,
    *,
    index: int,
    color: str | None,
    options: ChartOptions,
) -> None:
    """Format one series; `kind` is how it is drawn, `color` hex, `index` its position."""
    if kind in ("column", "bar", "area"):
        series.invertIfNegative = False
        _fill(series, color)
    elif kind == "bubble":
        series.invertIfNegative = False
        series.bubble3D = False
        _fill(series, color, alpha=_BUBBLE_ALPHA, index=index)
    elif kind == "line":
        _draw_line(series, spec, color)
        _marker(series, spec, "circle" if options.markers else "none", color)
        # Excel curves lines that carry no smooth flag.
        series.smooth = options.smooth
    elif kind == "radar":
        _draw_line(series, spec, color)
        _marker(series, spec, "none", color)
    elif kind == "scatter":
        _scatter(series, spec, options, color)
    labels = spec.data_labels or options.data_labels
    if labels:
        series.dLbls = DataLabelList(
            numFmt=labels.number_format,
            dLblPos=_LABEL_POSITIONS[labels.position] if labels.position else None,
            showLegendKey=False,
            showVal="value" in labels.show,
            showCatName="category" in labels.show,
            showSerName="series" in labels.show,
            showPercent="percent" in labels.show,
            showBubbleSize=False,
        )
        series.dLbls.spPr = look.no_fill()
        series.dLbls.txPr = look.text_properties(900)
    if spec.trendline:
        series.trendline = _trendline(spec.trendline, color, index)
    if spec.error_bars:
        bars = spec.error_bars
        series.errBars = XmlErrorBars(
            errDir=bars.axis,
            errBarType=bars.direction,
            errValType=_ERROR_VALUES[bars.kind],
            noEndCap=not bars.end_cap,
            val=bars.value,
            spPr=GraphicalProperties(ln=look.thin_line(65000, 35000)),
        )


def _trendline(trend: Trendline, color: str | None, index: int) -> XmlTrendline:
    line = LineProperties(
        w=19050,
        cap="rnd",
        solidFill=_hex(color) if color else look.theme_color(index),
        prstDash="sysDot",
    )
    label = TrendlineLabel(
        numFmt=NumFmt(formatCode="General", sourceLinked=False),
        spPr=look.no_fill(),
        txPr=look.text_properties(900),
    )
    return XmlTrendline(
        spPr=GraphicalProperties(ln=line),
        trendlineType=_TRENDLINES[trend.type],
        order=(trend.order or 2) if trend.type == "polynomial" else None,
        period=trend.period,
        dispRSqr=trend.r_squared,
        dispEq=trend.equation,
        trendlineLbl=label if trend.equation or trend.r_squared else None,
    )


def color_slices(series: Series, count: int, colors: list[str]) -> None:
    """Color each slice of a pie or doughnut, with Excel's white separators."""
    series.dPt = [
        DataPoint(
            idx=index,
            bubble3D=False,
            spPr=GraphicalProperties(
                solidFill=_hex(colors[index]) if index < len(colors) else look.theme_color(index),
                ln=LineProperties(w=19050, solidFill=look.scheme_color("lt1")),
            ),
        )
        for index in range(count)
    ]


def _hex(color: str) -> ColorChoice:
    return ColorChoice(srgbClr=color)


def _fill(series: Series, color: str | None, *, alpha: int | None = None, index: int = 0) -> None:
    """Fill a bar, area or bubble. Excel colors the series itself unless asked to."""
    if color:
        series.graphicalProperties = GraphicalProperties(solidFill=_hex(color))
    elif alpha:
        series.graphicalProperties = GraphicalProperties(
            solidFill=look.theme_color(index, alpha), ln=LineProperties(noFill=True)
        )


def _draw_line(series: Series, spec: SeriesSpec, color: str | None) -> None:
    if color or spec.line_width_pt:
        width = round(spec.line_width_pt * _EMU_PER_POINT) if spec.line_width_pt else None
        series.graphicalProperties = GraphicalProperties(
            ln=LineProperties(w=width, cap="rnd", solidFill=_hex(color) if color else None)
        )


def _marker(series: Series, spec: SeriesSpec, symbol: str, color: str | None) -> None:
    marker = Marker(symbol=spec.marker or symbol)
    if marker.symbol != "none":
        marker.size = spec.marker_size or 5
        if color:
            marker.graphicalProperties = GraphicalProperties(
                solidFill=_hex(color), ln=LineProperties(solidFill=_hex(color))
            )
    series.marker = marker


def _scatter(series: Series, spec: SeriesSpec, options: ChartOptions, color: str | None) -> None:
    """Excel's five scatter subtypes: which of markers and (smooth) lines are drawn."""
    style = options.scatter_style
    if style == "markers":
        series.graphicalProperties = GraphicalProperties(
            ln=LineProperties(w=19050, cap="rnd", noFill=True)
        )
    else:
        _draw_line(series, spec, color)
    series.smooth = style.startswith("smooth")
    has_markers = style in ("markers", "lines_markers", "smooth_markers")
    _marker(series, spec, "circle" if has_markers else "none", color)
