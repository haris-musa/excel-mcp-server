"""What a client can ask for in a chart: series, trendlines, error bars, axes and layout."""

from typing import Literal

from pydantic import Field

from excel_mcp.inputs import InputModel

ChartType = Literal[
    "column", "bar", "line", "area", "pie", "doughnut", "radar", "scatter", "bubble"
]
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


class DataLabels(InputModel):
    show: list[LabelContent] = Field(
        default=["value"], description="'percent' fits pie and doughnut charts only."
    )
    position: LabelPosition | None = Field(
        default=None,
        description="Default: Excel's. Valid positions depend on the chart type: column and "
        "bar charts take center, inside_end, inside_base, outside_end (not when stacked); line, "
        "scatter and bubble charts center, above, below, left, right; pie charts center, "
        "inside_end, outside_end, best_fit.",
    )
    number_format: str | None = Field(default=None, description="e.g. '0.0%' or '#,##0'.")


class Trendline(InputModel):
    type: TrendlineType
    order: int | None = Field(default=None, ge=2, le=6, description="Polynomial only. Default 2.")
    period: int | None = Field(default=None, ge=2, le=255, description="Moving average only.")
    equation: bool = Field(default=False, description="Show the equation on the chart.")
    r_squared: bool = Field(default=False, description="Show R² on the chart.")


class ErrorBars(InputModel):
    kind: ErrorBarKind
    value: float | None = Field(
        default=None,
        gt=0,
        description="Amount for 'fixed', percentage for 'percent', multiple of the standard "
        "deviation for 'std_dev'. Not used by 'std_error'.",
    )
    direction: Literal["both", "plus", "minus"] = "both"
    axis: Literal["y", "x"] = Field(default="y", description="'x' fits scatter and bubble.")
    end_cap: bool = True


class SeriesSpec(InputModel):
    values: str = Field(
        description="One row or column of values, e.g. 'B2:B13' or 'Data!B2:B13'. Without a "
        "sheet name, the chart's own sheet. For scatter and bubble charts these are the y values."
    )
    name: str | None = Field(
        default=None,
        description="Series name: text, or a sheet-qualified cell like 'Data!B1' that the name "
        "follows. Default: no name (Excel calls it 'Series1').",
    )
    categories: str | None = Field(
        default=None,
        description="Category labels, e.g. 'Data!A2:A13'; for scatter and bubble charts the x "
        "values. Default: the top-level `categories`.",
    )
    sizes: str | None = Field(default=None, description="Bubble sizes. Bubble charts only.")
    type: SeriesType | None = Field(
        default=None,
        description="Draw this series differently from chart_type, for a combo chart. "
        "Column, line and area charts only.",
    )
    secondary_axis: bool = Field(
        default=False,
        description="Plot on a second value axis (options.secondary_y_axis). Column, line, "
        "area and scatter charts only.",
    )
    color: str | None = Field(
        default=None,
        description="Hex, e.g. '#C00000': the fill of bars, areas and bubbles; the "
        "line of line and scatter series.",
    )
    line_width: float | None = Field(
        default=None, gt=0, le=20, description="Points. Line, scatter and radar series."
    )
    marker: MarkerStyle | None = Field(default=None, description="Line and scatter series.")
    marker_size: int | None = Field(default=None, ge=2, le=72)
    data_labels: DataLabels | None = Field(
        default=None, description="Overrides options.data_labels for this series."
    )
    trendline: Trendline | None = None
    error_bars: ErrorBars | None = None


class Axis(InputModel):
    title: str | None = None
    min: float | None = None
    max: float | None = None
    major_unit: float | None = Field(default=None, gt=0)
    log: bool = Field(default=False, description="Base-10 logarithmic scale.")
    reverse: bool = Field(
        default=False,
        description="Draw from the other end. In a bar chart the rows then run bottom-up, "
        "as Excel does by default.",
    )
    number_format: str | None = Field(default=None, description="e.g. '0%' or '#,##0'.")
    major_gridlines: bool | None = Field(
        default=None,
        description="Default: on for the main value axis (and the x axis of scatter and "
        "bubble charts), off otherwise.",
    )
    minor_gridlines: bool = False
    labels: TickLabels = Field(
        default="next_to_axis",
        description="Where the tick labels sit; for none at all, number_format ';;;'.",
    )


class ChartOptions(InputModel):
    title: str | None = None
    title_size: int | None = Field(default=None, ge=6, le=72, description="Points. Default 14.")
    width_cm: float = Field(default=15, gt=0, le=100)
    height_cm: float = Field(default=7.5, gt=0, le=100)
    legend: LegendPosition = Field(default="bottom", description="'none' hides it.")
    data_labels: DataLabels | None = Field(
        default=None, description="Labels on every series; a series' own data_labels win."
    )
    grouping: Grouping = Field(
        default="standard",
        description="'stacked' and 'percent_stacked' (categories sum to 100%) fit column, bar, "
        "line and area charts.",
    )
    colors: list[str] = Field(
        default_factory=list,
        description="Hex, e.g. ['#1F4E78', '#C00000']: one per series in order, or per slice "
        "in pie and doughnut charts. A series' own color wins.",
    )
    markers: bool = Field(default=False, description="Line charts only.")
    smooth: bool = Field(default=False, description="Line charts only.")
    scatter_style: ScatterStyle = Field(default="markers", description="Scatter charts only.")
    style: int | None = Field(
        default=None,
        ge=1,
        le=48,
        description="Excel 2007 chart style number, which sets the series colors (1 grayscale, "
        "2 colorful, 3-8 one accent color...). Default: Excel's own.",
    )
    plot_color: str | None = Field(default=None, description="Hex fill of the plot area.")
    x_axis: Axis = Field(
        default_factory=Axis,
        description="The category axis, which has no min, max, major_unit, log or "
        "number_format; for scatter and bubble charts, the x axis.",
    )
    y_axis: Axis = Field(default_factory=Axis, description="The value axis.")
    secondary_y_axis: Axis = Field(
        default_factory=Axis, description="Used by series with secondary_axis."
    )
