"""Rejecting chart requests that do not fit together, before anything is built."""

from excel_mcp.errors import InvalidArgumentError
from excel_mcp.operations.charts_data import Plot
from excel_mcp.operations.charts_options import (
    ROUND_TYPES,
    Axis,
    ChartOptions,
    ChartType,
    DataLabels,
    SeriesSpec,
)
from excel_mcp.operations.formatting import parse_color

MAX_SERIES = 255
_COMBINABLE = ("column", "line", "area")
_SECONDARY = ("column", "line", "area", "scatter")
_GROUPED = ("column", "bar", "line", "area")
_FITTED = ("column", "bar", "line", "area", "scatter", "bubble")
_MARKED = ("line", "scatter")
_LINED = ("line", "scatter", "radar")
_BARS = ("center", "inside_end", "inside_base", "outside_end")
_POINTS = ("center", "above", "below", "left", "right")
_SLICES = ("center", "inside_end", "outside_end", "best_fit")
_LABEL_POSITIONS = {
    "column": _BARS,
    "bar": _BARS,
    "line": _POINTS,
    "scatter": _POINTS,
    "bubble": _POINTS,
    "pie": _SLICES,
}
_VALUE_ONLY = ("min", "max", "major_unit", "number_format")


def check_chart(plots: list[Plot], chart_type: ChartType, options: ChartOptions) -> None:
    if len(plots) > MAX_SERIES:
        raise InvalidArgumentError(f"A chart has at most {MAX_SERIES} series, got {len(plots)}.")
    _check_options(options, chart_type)
    secondary = [plot.spec.secondary_axis for plot in plots]
    if all(secondary):
        raise InvalidArgumentError("At least one series must stay on the primary axis.")
    if not any(secondary) and options.secondary_y_axis != Axis():
        raise InvalidArgumentError(
            "secondary_y_axis is set, but no series has secondary_axis: true."
        )
    for number, plot in enumerate(plots, start=1):
        try:
            _check_series(plot, chart_type, options)
        except InvalidArgumentError as error:
            raise InvalidArgumentError(f"Series {number}: {error}") from None
    _check_colors(plots, chart_type, options)


def _check_options(options: ChartOptions, chart_type: ChartType) -> None:
    _require(options.grouping != "standard", "grouping", chart_type, _GROUPED)
    _require(options.markers, "markers", chart_type, ("line",))
    _require(options.smooth, "smooth", chart_type, ("line",))
    _require(options.scatter_style != "markers", "scatter_style", chart_type, ("scatter",))
    axes = (options.x_axis, options.y_axis, options.secondary_y_axis)
    if chart_type in ROUND_TYPES and any(axis != Axis() for axis in axes):
        raise InvalidArgumentError(f"{chart_type} charts have no axes to set.")
    for name, axis in (
        ("x_axis", options.x_axis),
        ("y_axis", options.y_axis),
        ("secondary_y_axis", options.secondary_y_axis),
    ):
        _check_axis(name, axis, name == "x_axis" and chart_type not in ("scatter", "bubble"))


def _check_axis(name: str, axis: Axis, category: bool) -> None:
    if category:
        used = [field for field in _VALUE_ONLY if getattr(axis, field) is not None] + (
            ["log"] if axis.log else []
        )
        if used:
            raise InvalidArgumentError(
                f"{name} is a category axis, which has no {', '.join(used)}. Scale it on a "
                "value axis (y_axis)."
            )
    if axis.min is not None and axis.max is not None and axis.min >= axis.max:
        raise InvalidArgumentError(f"{name}.min ({axis.min}) must be below its max ({axis.max}).")
    if axis.log and axis.min is not None and axis.min <= 0:
        raise InvalidArgumentError(f"{name}: a logarithmic axis needs a min above zero (or none).")


def _require(used: bool, name: str, chart_type: ChartType, allowed: tuple[str, ...]) -> None:
    if used and chart_type not in allowed:
        raise InvalidArgumentError(
            f"{name} does not apply to {chart_type} charts; it works with {', '.join(allowed)}."
        )


def _check_series(plot: Plot, chart_type: ChartType, options: ChartOptions) -> None:
    spec = plot.spec
    if spec.type is not None and chart_type not in _COMBINABLE:
        raise InvalidArgumentError(
            f"type makes a combo chart, which needs chart_type column, line or area, not "
            f"{chart_type}."
        )
    kind = spec.type or chart_type
    _require(spec.secondary_axis, "secondary_axis", kind, _SECONDARY)
    _require(spec.marker is not None or spec.marker_size is not None, "marker", kind, _MARKED)
    _require(spec.line_width is not None, "line_width", kind, _LINED)
    if (plot.sizes is not None) != (chart_type == "bubble"):
        raise InvalidArgumentError(
            "sizes are required for bubble charts, and only for them."
            if chart_type == "bubble"
            else f"sizes apply to bubble charts only, not {chart_type}."
        )
    stacked = options.grouping != "standard" and kind == chart_type
    if spec.trendline:
        _require(True, "trendline", kind, _FITTED)
        if stacked:
            raise InvalidArgumentError("Excel cannot fit a trendline to a stacked series.")
        _check_trendline(spec)
    if spec.error_bars:
        _require(True, "error_bars", kind, _FITTED)
        _check_error_bars(spec, kind)
    labels = spec.data_labels or options.data_labels
    if labels:
        _check_labels(labels, kind, stacked)


def _check_trendline(spec: SeriesSpec) -> None:
    trend = spec.trendline
    assert trend
    if trend.order is not None and trend.type != "polynomial":
        raise InvalidArgumentError("trendline.order applies to polynomial trendlines only.")
    if trend.period is not None and trend.type != "moving_average":
        raise InvalidArgumentError("trendline.period applies to moving_average trendlines only.")
    if trend.type == "moving_average" and trend.period is None:
        raise InvalidArgumentError("A moving_average trendline needs a period, e.g. 3.")
    if trend.type == "moving_average" and (trend.equation or trend.r_squared):
        raise InvalidArgumentError("A moving average has no equation or R².")


def _check_error_bars(spec: SeriesSpec, kind: str) -> None:
    bars = spec.error_bars
    assert bars
    if bars.axis == "x" and kind not in ("scatter", "bubble"):
        raise InvalidArgumentError("x error bars fit scatter and bubble charts only.")
    if bars.kind == "std_error" and bars.value is not None:
        raise InvalidArgumentError("std_error error bars take no value.")
    if bars.kind != "std_error" and bars.value is None:
        raise InvalidArgumentError(f"{bars.kind} error bars need a value.")


def _check_labels(labels: DataLabels, kind: str, stacked: bool) -> None:
    if "percent" in labels.show and kind not in ROUND_TYPES:
        raise InvalidArgumentError("Percent labels fit pie and doughnut charts only.")
    if labels.position is None:
        return
    allowed = _LABEL_POSITIONS.get(kind, ())
    if stacked:
        allowed = tuple(position for position in allowed if position != "outside_end")
    if labels.position not in allowed:
        valid = ", ".join(allowed) or "none: the labels sit where Excel puts them"
        raise InvalidArgumentError(
            f"Label position {labels.position!r} does not fit a {'stacked ' * stacked}{kind} "
            f"series. Valid positions: {valid}."
        )


def _check_colors(plots: list[Plot], chart_type: ChartType, options: ChartOptions) -> None:
    slices = plots[0].points
    count = slices if chart_type in ROUND_TYPES else len(plots)
    if len(options.colors) > count:
        what = "slices" if chart_type in ROUND_TYPES else "series"
        raise InvalidArgumentError(f"{len(options.colors)} colors given for {count} {what}.")
    for color in [*options.colors, options.plot_color, *(plot.spec.color for plot in plots)]:
        if color is not None:
            parse_color(color)
