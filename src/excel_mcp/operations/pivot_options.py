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
    field: str = Field(description="Header of the column to summarize, or a calculated field.")
    function: Function = Field(
        default="sum", description="'count' counts non-empty cells; the others need numbers."
    )
    number_format: str | None = Field(
        default=None,
        max_length=255,
        description="Excel code such as '#,##0.00'. Default: the source's format.",
    )
    show_as: ShowAs | None = Field(
        default=None,
        description="Show each value relative to others, as 'Show Values As' does. "
        "Needs base_field for percent_of_parent, difference_from, percent_difference_from, "
        "percent_of, running_total, percent_running_total and the rank types.",
    )
    base_field: str | None = Field(
        default=None, description="A rows or columns field that show_as compares along."
    )
    base_item: str | None = Field(
        default=None,
        description="With difference_from, percent_difference_from or percent_of: the item to "
        "compare with. Default: the previous item.",
    )


class NumberGroup(InputModel):
    by: float = Field(gt=0, description="Size of each range.")
    start: float | None = Field(default=None, description="Default: the smallest value.")
    end: float | None = Field(default=None, description="Default: the largest value.")


class PivotField(InputModel):
    field: str = Field(description="A header used in rows, columns or filters.")
    show_items: list[str] = Field(
        default=[], description="Show only these items (as the table labels them). Default: all."
    )
    sort: Literal["ascending", "descending"] = Field(
        default="ascending", description="Order of the items."
    )
    sort_by: str | None = Field(
        default=None, description="Sort by this values field's total (e.g. 'Sum of Units')."
    )
    group_dates: list[DateUnit] = Field(
        default=[],
        description="Group a date field, e.g. ['years', 'months']; each unit becomes a "
        "level, largest first.",
    )
    group_numbers: NumberGroup | None = Field(
        default=None, description="Group a number field into ranges."
    )


class CalculatedField(InputModel):
    name: str = Field(description="Name to use in values, e.g. 'Revenue'.")
    formula: str = Field(
        max_length=8_192,
        description="Over source field names (quote names with spaces: 'Unit Price'), "
        "e.g. 'Units*Price'. Applied to the sums, as in Excel.",
    )


@dataclass(frozen=True)
class DatePeriod:
    """Records whose date in ``field`` lies from ``start`` to ``end``, both included."""

    field: str
    start: dt.datetime
    end: dt.datetime
