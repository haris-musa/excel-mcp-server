"""What a group of sparklines looks like: Insert > Sparklines and the Sparkline tab of Excel."""

from typing import Literal

from pydantic import Field

from excel_mcp.errors import InvalidArgumentError
from excel_mcp.inputs import InputModel

SparklineType = Literal["line", "column", "win_loss"]
Point = Literal["markers", "high", "low", "first", "last", "negative"]
AxisScale = Literal["individual", "same"] | float
LINE_WEIGHT = 0.75


class SparklineColors(InputModel):
    """Hex RGB colors such as '376092'."""

    series: str = "376092"
    negative: str = "D00000"
    axis: str = "000000"
    markers: str = "D00000"
    first: str = "D00000"
    last: str = "D00000"
    high: str = "D00000"
    low: str = "D00000"


class SparklineStyle(InputModel):
    type: SparklineType = "line"
    colors: SparklineColors = Field(default_factory=SparklineColors)
    show: list[Point] = Field(
        default_factory=list,
        description="Points to emphasize in their colors. markers is for line only; "
        "win_loss shows losses with negative.",
    )
    show_axis: bool = Field(default=False, description="Draw the horizontal axis (0 or dates).")
    axis_min: AxisScale = Field(
        default="individual",
        description="Lowest value of the vertical axis: each sparkline's own, the lowest of "
        "the group ('same'), or a number.",
    )
    axis_max: AxisScale = Field(default="individual", description="Like axis_min.")
    right_to_left: bool = Field(default=False, description="Plot the data from right to left.")
    dates: str | None = Field(
        default=None,
        description="Range of dates, one per data point, to plot on a date axis, "
        "e.g. 'Data!B1:F1'.",
    )
    empty_cells: Literal["gap", "zero", "connect"] = Field(
        default="gap", description="Show empty cells as gaps or zeros, or connect the points."
    )
    hidden: bool = Field(default=False, description="Plot data in hidden rows and columns.")
    line_weight: float = Field(default=LINE_WEIGHT, gt=0, le=100, description="Line only, points.")


class SparklineInfo(SparklineStyle):
    sparklines: dict[str, str] = Field(description="Cell to the data it plots.")


def check_style(style: SparklineStyle) -> None:
    if style.type != "line" and (
        "markers" in style.show or "line_weight" in style.model_fields_set
    ):
        raise InvalidArgumentError(
            f"markers and line_weight are for line sparklines, not {style.type}."
        )
