"""What a client can ask of a PivotTable beyond its rows, columns and values."""

import datetime as dt
from dataclasses import dataclass
from typing import Literal

from pydantic import Field

from excel_mcp.inputs import InputModel

Function = Literal["sum", "count", "average", "min", "max"]
ShowAs = Literal[
    "percent_of_total",
    "percent_of_row",
    "percent_of_column",
    "percent_of_parent_row",
    "percent_of_parent_column",
    "percent_of_parent",
    "difference_from",
    "percent_difference_from",
    "percent_of",
    "running_total",
    "percent_running_total",
    "rank_ascending",
    "rank_descending",
]
DateUnit = Literal["years", "quarters", "months", "days"]
Layout = Literal["compact", "outline", "tabular"]
ValuesIn = Literal["columns", "rows"]


class PivotValue(InputModel):
    field: str = Field(description="Column header or calculated field.")
    function: Function = Field(
        default="sum", description="'count' counts non-empty cells; others need numbers."
    )
    number_format: str | None = Field(
        default=None,
        max_length=255,
        description="Excel code. Default: the source's.",
    )
    show_as: ShowAs | None = Field(
        default=None,
        description="'Show Values As'. Needs base_field, except for percent_of_total, "
        "percent_of_row, percent_of_column, percent_of_parent_row and _column.",
    )
    base_field: str | None = Field(
        default=None, description="Row or column field that show_as runs along."
    )
    base_item: str | None = Field(
        default=None,
        description="difference_from, percent_difference_from, percent_of: item to compare "
        "with. Default: the previous.",
    )


class NumberGroup(InputModel):
    by: float = Field(gt=0, description="Size of each range.")
    start: float | None = Field(default=None, description="Default: the smallest value.")
    end: float | None = Field(default=None, description="Default: the largest value.")


class PivotField(InputModel):
    field: str = Field(description="A header used in rows, columns or filters.")
    show_items: list[str] = Field(default=[], description="Show only these items. Default: all.")
    sort: Literal["ascending", "descending"] = Field(default="ascending")
    sort_by: str | None = Field(
        default=None, description="Sort by this value field's total, e.g. 'Sum of Units'."
    )
    group_dates: list[DateUnit] = Field(
        default=[],
        description="Group a date field, e.g. ['years', 'months'] (largest first).",
    )
    group_numbers: NumberGroup | None = Field(
        default=None, description="Group numbers into ranges."
    )


class CalculatedField(InputModel):
    name: str = Field(description="e.g. 'Revenue'.")
    formula: str = Field(
        max_length=8_192,
        description="Over source field names ('Unit Price' quoted), e.g. 'Units*Price'; "
        "applied to the sums, as in Excel.",
    )


@dataclass(frozen=True)
class DatePeriod:
    """Records whose date in ``field`` lies from ``start`` to ``end``, both included."""

    field: str
    start: dt.datetime
    end: dt.datetime
