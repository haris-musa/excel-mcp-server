"""What a client can ask for in a chart: series, trendlines, error bars, axes and layout."""

from typing import Literal

from pydantic import Field

from excel_mcp.inputs import InputModel

ChartType = Literal[
    "column", "bar", "line", "area", "pie", "doughnut", "radar", "scatter", "bubble",
    "waterfall", "histogram", "pareto", "box_whisker", "treemap", "sunburst", "funnel",
]  # fmt: skip
SeriesType = Literal["column", "line", "area"]
Grouping = Literal["standard", "stacked", "percent_stacked"]
LegendPosition = Literal["right", "left", "top", "bottom", "none"]
ScatterStyle = Literal["markers", "lines_markers", "lines", "smooth_markers", "smooth"]
MarkerStyle = Literal[
    "circle", "square", "diamond", "triangle", "x", "star", "dash", "dot", "plus", "none"
]
LabelPosition = Literal[
    "center", "inside_end", "inside_base", "outside_end", "above", "below", "left", "right",
    "best_fit",
]  # fmt: skip
LabelContent = Literal["value", "percent", "category", "series"]
TrendlineType = Literal[
    "linear", "exponential", "logarithmic", "polynomial", "power", "moving_average"
]
ErrorBarKind = Literal["fixed", "percent", "std_dev", "std_error"]
TickLabels = Literal["next_to_axis", "low", "high"]

ROUND_TYPES = ("pie", "doughnut")
MODERN_TYPES = (
    "waterfall", "histogram", "pareto", "box_whisker", "treemap", "sunburst", "funnel"
)  # fmt: skip
ParentLabels = Literal["overlapping", "banner", "none"]
Quartiles = Literal["exclusive", "inclusive"]


class DataLabels(InputModel):
    show: list[LabelContent] = Field(
        default=["value"],
        description="'percent': pie and doughnut only. Empty: none.",
    )
    position: LabelPosition | None = Field(
        default=None,
        description="Column, bar (not stacked: outside_end), waterfall, histogram, pareto: "
        "center, inside_end, inside_base, outside_end. Line, scatter, bubble: center, above, "
        "below, left, right. Pie: center, inside_end, outside_end, best_fit.",
    )
    number_format: str | None = None


class Trendline(InputModel):
    type: TrendlineType
    order: int | None = Field(default=None, ge=2, le=6, description="Polynomial. Default 2.")
    period: int | None = Field(default=None, ge=2, le=255, description="Moving average.")
    equation: bool = False
    r_squared: bool = False


class ErrorBars(InputModel):
    kind: ErrorBarKind
    value: float | None = Field(
        default=None,
        gt=0,
        description="Amount, percentage or standard deviations; not for 'std_error'.",
    )
    direction: Literal["both", "plus", "minus"] = "both"
    axis: Literal["y", "x"] = Field(default="y", description="'x': scatter and bubble.")
    end_cap: bool = True


class SeriesSpec(InputModel):
    values: str = Field(
        description="One row or column, e.g. 'Data!B2:B13' (no sheet: the chart's). Scatter "
        "and bubble: y values."
    )
    name: str | None = Field(
        default=None,
        description="Text, or a sheet-qualified cell like 'Data!B1'.",
    )
    categories: str | None = Field(
        default=None,
        description="Default: the top-level `categories`. Scatter and bubble: x values.",
    )
    sizes: str | None = Field(default=None, description="Bubble charts.")
    type: SeriesType | None = Field(
        default=None,
        description="Draw it differently from chart_type (combo chart).",
    )
    secondary_axis: bool = Field(
        default=False, description="Use options.secondary_y_axis, a second value axis."
    )
    color: str | None = Field(
        default=None,
        description="Hex, e.g. '#C00000': fill, or the line of line and scatter series.",
    )
    line_width_pt: float | None = Field(default=None, gt=0, le=20)
    marker: MarkerStyle | None = None
    marker_size: int | None = Field(default=None, ge=2, le=72)
    data_labels: DataLabels | None = Field(
        default=None, description="Overrides options.data_labels."
    )
    trendline: Trendline | None = None
    error_bars: ErrorBars | None = None


class Axis(InputModel):
    title: str | None = None
    min: float | None = None
    max: float | None = None
    major_unit: float | None = Field(default=None, gt=0)
    log: bool = False
    reverse: bool = False
    number_format: str | None = None
    major_gridlines: bool | None = Field(
        default=None, description="Default: on for the value axis, off for categories."
    )
    minor_gridlines: bool = False
    labels: TickLabels = Field(
        default="next_to_axis",
        description="Tick label position; to hide them, number_format ';;;'.",
    )


def default_legend(chart_type: str, series: int) -> LegendPosition:
    """Where Excel puts the legend of a new chart, as found by creating each type in Excel."""
    if chart_type in ("waterfall", "treemap"):
        return "top"
    if chart_type in ROUND_TYPES:
        return "bottom"
    if series == 1 or chart_type in MODERN_TYPES or chart_type == "bubble":
        return "none"
    return "top" if chart_type == "radar" else "bottom"


class Bins(InputModel):
    width: float | None = Field(default=None, gt=0, description="Values per bin.")
    count: int | None = Field(default=None, ge=1, le=1000)
    underflow: float | None = Field(
        default=None, description="One bin for values at or below this."
    )
    overflow: float | None = Field(default=None, description="One bin for values above this.")


class BoxPlot(InputModel):
    quartiles: Quartiles = "exclusive"
    mean_marker: bool = True
    mean_line: bool = False
    inner_points: bool = False
    outliers: bool = True


class ChartOptions(InputModel):
    title: str | None = None
    title_size: int | None = Field(default=None, ge=6, le=72, description="Points. Default 14.")
    width_cm: float = Field(default=15, gt=0, le=100)
    height_cm: float = Field(default=7.5, gt=0, le=100)
    legend: LegendPosition | None = Field(
        default=None, description="Default: Excel's for the type."
    )
    data_labels: DataLabels | None = Field(
        default=None, description="All series; a series' own data_labels win."
    )
    grouping: Grouping = Field(
        default="standard",
        description="Stacked: column, bar, line, area.",
    )
    colors: list[str] = Field(
        default_factory=list,
        description="Hex, per series or pie slice in order. A series' color wins.",
    )
    markers: bool = Field(default=False, description="Line charts.")
    smooth: bool = Field(default=False, description="Line charts.")
    scatter_style: ScatterStyle = "markers"
    style: int | None = Field(
        default=None,
        ge=1,
        le=48,
        description="Excel 2007 chart style (1 grayscale, 2 colorful, 3-8 one accent color).",
    )
    plot_color: str | None = Field(default=None, description="Hex.")
    x_axis: Axis = Field(
        default_factory=Axis,
        description="Category axis (no min, max, major_unit, log, number_format); x values "
        "in scatter and bubble.",
    )
    y_axis: Axis = Field(default_factory=Axis)
    secondary_y_axis: Axis = Field(
        default_factory=Axis,
        description="For series with secondary_axis; pareto: the percentage axis (0 to 1).",
    )
    totals: list[int] = Field(
        default_factory=list,
        description="Waterfall: 1-based positions of total points.",
    )
    connector_lines: bool = Field(default=True, description="Waterfall.")
    bins: Bins = Field(
        default_factory=Bins,
        description="Histogram, or pareto of numbers. Default: automatic.",
    )
    box: BoxPlot = Field(default_factory=BoxPlot)
    parent_labels: ParentLabels = Field(default="overlapping", description="Treemap group labels.")
