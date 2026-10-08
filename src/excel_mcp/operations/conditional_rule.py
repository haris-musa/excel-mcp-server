"""The conditional format rule a client describes."""

from typing import Literal

from pydantic import Field

from excel_mcp.inputs import InputModel
from excel_mcp.operations.conditional_kinds import IconSetName, Period, RuleType, ScaleType
from excel_mcp.operations.rules import Operator


class ConditionalFormat(InputModel):
    """Only the fields named for the ``type`` apply; others are rejected."""

    type: RuleType = Field(
        description="top/bottom: highest or lowest `count` values (or percent). "
        "above_average/below_average: compared with the range's average. "
        "date: dates in a `period`."
    )
    colors: list[str] | None = Field(
        default=None, description="color_scale: 2 or 3, lowest to highest. data_bar: 1."
    )
    operator: Operator | None = Field(default=None, description="cell_value.")
    values: list[str] | None = Field(
        default=None,
        description="cell_value: 1, or 2 for between/notBetween. Numbers, quoted text such "
        "as '\"Done\"', or formulas.",
    )
    formula: str | None = Field(
        default=None,
        description="formula: true for highlighted cells, written for the range's top-left "
        "cell, e.g. '=$C2>100'.",
    )
    count: int | None = Field(
        default=None, ge=1, le=1000, description="top, bottom: how many (1-100 if percent)."
    )
    percent: bool = Field(default=False, description="top, bottom: `count` is a percentage.")
    std_dev: int | None = Field(
        default=None, ge=1, le=3, description="above/below_average: standard deviations."
    )
    include_equal: bool = Field(default=False, description="above/below_average.")
    text: str | None = Field(default=None, max_length=255, description="contains_text etc.")
    period: Period | None = Field(default=None, description="date.")
    icon_set: IconSetName | None = Field(default=None, description="icon_set.")
    thresholds: list[float] | None = Field(
        default=None,
        description="icon_set: where icons 2..n start, lowest first (n-1 values). "
        "Default: equal shares as in Excel.",
    )
    threshold_type: Literal["percent", "number", "percentile"] = Field(
        default="percent", description="icon_set: what `thresholds` mean."
    )
    icons: list[str] | None = Field(
        default=None,
        description="icon_set: a custom icon for each of the n ranges, lowest first: "
        "'3Arrows:1' is the lowest icon of that set (1) up to its highest (3, 4 or 5), "
        "'none' shows no icon. Sets can be mixed.",
    )
    reverse: bool = Field(default=False, description="icon_set: reverse the icon order.")
    hide_values: bool = Field(default=False, description="icon_set, data_bar: hide the values.")
    bar_fill: Literal["gradient", "solid"] = Field(default="gradient", description="data_bar.")
    border_color: str | None = Field(default=None, description="data_bar: border of the bars.")
    negative_color: str | None = Field(
        default=None, description="data_bar: fill of negative bars. Default: as `colors`."
    )
    negative_border_color: str | None = Field(
        default=None, description="data_bar: border of negative bars. Default: border_color."
    )
    axis: Literal["automatic", "middle", "none"] = Field(
        default="none",
        description="data_bar: where negative bars start. automatic: in proportion to the "
        "values; middle: in the middle of the cell.",
    )
    axis_color: str | None = Field(default=None, description="data_bar: default black.")
    bar_direction: Literal["context", "left_to_right", "right_to_left"] = Field(
        default="context", description="data_bar: context follows the text direction."
    )
    min_type: ScaleType = Field(
        default="lowest", description="data_bar: what the shortest bar stands for."
    )
    min_value: float | str | None = Field(
        default=None,
        description="data_bar: with min_type number, percent, percentile or formula "
        "(a formula such as '=$G$1').",
    )
    max_type: ScaleType = Field(
        default="highest", description="data_bar: what the longest bar stands for."
    )
    max_value: float | str | None = Field(default=None, description="data_bar: like min_value.")
    fill_color: str | None = Field(default=None, description="Every type but the scales/icons.")
    font_color: str | None = Field(default=None, description="Like fill_color.")
    stop_if_true: bool = Field(default=False, description="Skip lower-priority rules if met.")
    priority: int | None = Field(
        default=None,
        ge=1,
        description="1 is evaluated first; rules at or below it move down. Default: last.",
    )
