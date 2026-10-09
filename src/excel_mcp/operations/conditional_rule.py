"""The conditional format rule a client describes."""

from typing import Literal

from pydantic import Field

from excel_mcp.inputs import InputModel
from excel_mcp.operations.conditional_kinds import IconSetName, Period, RuleType, ScaleType
from excel_mcp.operations.rules import Operator


class ConditionalFormat(InputModel):
    """Fields apply per `type`."""

    type: RuleType = Field(
        description="top/bottom: `count` highest or lowest. date: dates in a `period`."
    )
    colors: list[str] | None = Field(
        default=None, description="color_scale: 2 or 3, lowest first. data_bar: 1."
    )
    operator: Operator | None = Field(default=None, description="cell_value.")
    values: list[str] | None = Field(
        default=None,
        description="cell_value: 1, or 2 for between. Numbers, quoted text '\"Done\"' or formulas.",
    )
    formula: str | None = Field(
        default=None,
        description="formula: for the range's top-left cell, e.g. '=$C2>100'.",
    )
    count: int | None = Field(
        default=None, ge=1, le=1000, description="top, bottom (1-100 if percent)."
    )
    percent: bool = Field(default=False, description="top, bottom.")
    std_dev: int | None = Field(default=None, ge=1, le=3, description="above/below_average.")
    include_equal: bool = Field(default=False, description="above/below_average.")
    text: str | None = Field(default=None, max_length=255, description="contains_text etc.")
    period: Period | None = Field(default=None, description="date.")
    icon_set: IconSetName | None = Field(default=None, description="icon_set.")
    thresholds: list[float] | None = Field(
        default=None,
        description="icon_set: where icons 2..n start, lowest first. Default: equal shares.",
    )
    threshold_type: Literal["percent", "number", "percentile"] = Field(
        default="percent", description="icon_set."
    )
    icons: list[str] | None = Field(
        default=None,
        description="icon_set: a custom icon per range, lowest first, e.g. '3Arrows:1' "
        "(the set's lowest icon) or 'none'. Sets can be mixed.",
    )
    reverse: bool = Field(default=False, description="icon_set.")
    hide_values: bool = Field(default=False, description="icon_set, data_bar.")
    bar_fill: Literal["gradient", "solid"] = Field(default="gradient", description="data_bar.")
    border_color: str | None = Field(default=None, description="data_bar.")
    negative_color: str | None = Field(default=None, description="data_bar. Default: as `colors`.")
    negative_border_color: str | None = Field(
        default=None, description="data_bar. Default: border_color."
    )
    axis: Literal["automatic", "middle", "none"] = Field(
        default="none",
        description="data_bar: where negative bars start.",
    )
    axis_color: str | None = Field(default=None, description="data_bar.")
    bar_direction: Literal["context", "left_to_right", "right_to_left"] = Field(
        default="context", description="data_bar."
    )
    min_type: ScaleType = Field(default="lowest", description="data_bar: the shortest bar.")
    min_value: float | str | None = Field(
        default=None,
        description="data_bar: for min_type number, percent, percentile or formula.",
    )
    max_type: ScaleType = Field(default="highest", description="data_bar: the longest bar.")
    max_value: float | str | None = Field(default=None, description="data_bar.")
    fill_color: str | None = Field(default=None, description="Not for scales and icons.")
    font_color: str | None = Field(default=None, description="Like fill_color.")
    stop_if_true: bool = Field(default=False, description="Skip later rules if met.")
    priority: int | None = Field(
        default=None,
        ge=1,
        description="1 is first; rules at or below move down. Default: last.",
    )
