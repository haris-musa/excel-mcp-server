"""Chart options: what the client can ask for, and how each option is applied."""

from typing import Literal

from pydantic import BaseModel, Field

from excel_mcp.errors import InvalidArgumentError
from excel_mcp.operations.formatting import parse_color

ChartType = Literal["column", "bar", "line", "area", "pie", "scatter", "doughnut", "radar"]
Grouping = Literal["standard", "stacked", "percent_stacked"]

_GROUPED = ("column", "bar", "line", "area")
ROUND_TYPES = ("pie", "doughnut")


class ChartOptions(BaseModel):
    """Chart titles, size, legend, labels, colors and axes."""

    title: str | None = Field(default=None, description="Chart title.")
    x_axis_title: str | None = Field(default=None, description="Horizontal axis title.")
    y_axis_title: str | None = Field(default=None, description="Vertical axis title.")
    width_cm: float = Field(default=15, gt=0, le=100, description="Chart width in cm.")
    height_cm: float = Field(default=7.5, gt=0, le=100, description="Chart height in cm.")
    show_legend: bool = Field(default=True, description="Show the series legend.")
    legend_position: Literal["right", "left", "top", "bottom"] | None = Field(
        default=None,
        description="Where the legend sits. Default: right. Needs show_legend true.",
    )
    data_labels: bool = Field(
        default=False, description="Print each value on its bar, point or slice."
    )
    grouping: Grouping | None = Field(
        default=None,
        description="How series combine: 'standard' (side by side for column/bar), 'stacked' "
        "or 'percent_stacked' (each category sums to 100%). For column, bar, line and area "
        "charts only. Default: standard.",
    )
    colors: list[str] | None = Field(
        default=None,
        description="Hex colors such as ['#1F4E78', '#C00000'], one per series in the order of "
        "the data columns (fewer colors leave the remaining series on the default palette). "
        "For pie and doughnut charts, one color per slice instead.",
    )
    markers: bool | None = Field(
        default=None,
        description="Show (true) or hide (false) point markers. Line and scatter charts only. "
        "Default: the chart type's own style.",
    )
    smooth: bool | None = Field(
        default=None,
        description="Draw curved (true) or straight (false) lines. Line and scatter charts "
        "only. Default: the chart type's own style.",
    )
    y_axis_min: float | None = Field(
        default=None, description="Lowest value on the (primary) vertical axis. Default: automatic."
    )
    y_axis_max: float | None = Field(
        default=None,
        description="Highest value on the (primary) vertical axis. Default: automatic.",
    )
    y_axis_number_format: str | None = Field(
        default=None,
        description="Excel number format for the (primary) vertical axis labels, e.g. '0%' "
        "or '#,##0'. Default: taken from the data.",
    )
    secondary_line_columns: list[str] | None = Field(
        default=None,
        description="Header names of data columns to draw as lines on a secondary vertical "
        "axis on the right, combined with the other data columns drawn as columns (a combo "
        "chart). For column charts only; at least one data column must stay as columns.",
    )


def check_options(options: ChartOptions, chart_type: ChartType) -> None:
    """Reject options that do not apply to the chart type, before anything is built."""

    def require(name: str, allowed: tuple[str, ...]) -> None:
        if getattr(options, name) is not None and chart_type not in allowed:
            raise InvalidArgumentError(
                f"{name} does not apply to {chart_type} charts; it works with {', '.join(allowed)}."
            )

    if options.legend_position is not None and not options.show_legend:
        raise InvalidArgumentError("legend_position needs show_legend to be true.")
    require("grouping", _GROUPED)
    require("markers", ("line", "scatter"))
    require("smooth", ("line", "scatter"))
    require("secondary_line_columns", ("column",))
    for name in ("y_axis_min", "y_axis_max", "y_axis_number_format"):
        if getattr(options, name) is not None and chart_type in ROUND_TYPES:
            raise InvalidArgumentError(f"{name} does not apply to {chart_type} charts: no axes.")
    low, high = options.y_axis_min, options.y_axis_max
    if low is not None and high is not None and low >= high:
        raise InvalidArgumentError(f"y_axis_min ({low}) must be below y_axis_max ({high}).")
    if options.secondary_line_columns is not None and not options.secondary_line_columns:
        raise InvalidArgumentError("secondary_line_columns must name at least one column.")


def parse_colors(options: ChartOptions, chart_type: ChartType, slots: int) -> list[str]:
    """Colors as 6-digit hex, checked against the number of series (or pie slices)."""
    colors = options.colors or []
    if len(colors) > slots:
        what = "slices" if chart_type in ROUND_TYPES else "series"
        raise InvalidArgumentError(f"{len(colors)} colors given for {slots} {what}.")
    return [parse_color(color)[2:] for color in colors]
