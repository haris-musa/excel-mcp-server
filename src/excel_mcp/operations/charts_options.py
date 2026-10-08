"""Chart options: what the client can ask for, and which chart types each option fits."""

from typing import Literal

from pydantic import BaseModel, Field

from excel_mcp.errors import InvalidArgumentError
from excel_mcp.operations.formatting import parse_color

ChartType = Literal["column", "bar", "line", "area", "pie", "scatter", "doughnut", "radar"]
Grouping = Literal["standard", "stacked", "percent_stacked"]
LegendPosition = Literal["right", "left", "top", "bottom", "none"]

ROUND_TYPES = ("pie", "doughnut")
_GROUPED = ("column", "bar", "line", "area")


class ChartOptions(BaseModel):
    """Chart titles, size, legend, labels, colors and axes."""

    title: str | None = Field(default=None, description="Chart title.")
    x_axis_title: str | None = Field(default=None, description="Horizontal axis title.")
    y_axis_title: str | None = Field(default=None, description="Vertical axis title.")
    width_cm: float = Field(default=15, gt=0, le=100, description="Chart width in cm.")
    height_cm: float = Field(default=7.5, gt=0, le=100, description="Chart height in cm.")
    legend: LegendPosition = Field(
        default="right", description="Where the series legend sits, or 'none' to hide it."
    )
    data_labels: bool = Field(
        default=False, description="Print each value on its bar, point or slice."
    )
    grouping: Grouping = Field(
        default="standard",
        description="How series combine: 'standard' (side by side), 'stacked', or "
        "'percent_stacked' (each category sums to 100%). Stacking works for column, bar, "
        "line and area charts.",
    )
    colors: list[str] = Field(
        default_factory=list,
        description="Hex colors such as ['#1F4E78', '#C00000'], one per series in the order "
        "of the data columns; series without a color use Excel's palette. For pie and "
        "doughnut charts, one color per slice.",
    )
    markers: bool = Field(default=False, description="Mark each point. Line charts only.")
    smooth: bool = Field(default=False, description="Draw curved lines. Line charts only.")
    y_axis_min: float | None = Field(
        default=None, description="Lowest value on the vertical axis. Default: automatic."
    )
    y_axis_max: float | None = Field(
        default=None, description="Highest value on the vertical axis. Default: automatic."
    )
    y_axis_number_format: str | None = Field(
        default=None,
        description="Excel number format for the vertical axis labels, e.g. '0%' or '#,##0'. "
        "Default: the data's format.",
    )
    secondary_line_columns: list[str] = Field(
        default_factory=list,
        description="Header names of data columns to draw as lines on a second vertical axis "
        "on the right, while the other columns stay as columns (a combo chart). Column "
        "charts only.",
    )


def check_options(options: ChartOptions, chart_type: ChartType) -> None:
    """Reject options that do not fit the chart type, before anything is built."""
    _require(options.grouping != "standard", "grouping", chart_type, _GROUPED)
    _require(options.markers, "markers", chart_type, ("line",))
    _require(options.smooth, "smooth", chart_type, ("line",))
    _require(
        bool(options.secondary_line_columns), "secondary_line_columns", chart_type, ("column",)
    )
    axis_options = (options.y_axis_min, options.y_axis_max, options.y_axis_number_format)
    if chart_type in ROUND_TYPES and any(value is not None for value in axis_options):
        raise InvalidArgumentError(f"{chart_type} charts have no axes to set.")
    low, high = options.y_axis_min, options.y_axis_max
    if low is not None and high is not None and low >= high:
        raise InvalidArgumentError(f"y_axis_min ({low}) must be below y_axis_max ({high}).")


def _require(used: bool, name: str, chart_type: ChartType, allowed: tuple[str, ...]) -> None:
    if used and chart_type not in allowed:
        raise InvalidArgumentError(
            f"{name} does not apply to {chart_type} charts; it works with {', '.join(allowed)}."
        )


def parse_colors(options: ChartOptions, chart_type: ChartType, slots: int) -> list[str]:
    """Colors as 6-digit hex, checked against the number of series (or slices)."""
    if len(options.colors) > slots:
        what = "slices" if chart_type in ROUND_TYPES else "series"
        raise InvalidArgumentError(f"{len(options.colors)} colors given for {slots} {what}.")
    return [parse_color(color)[2:] for color in options.colors]
