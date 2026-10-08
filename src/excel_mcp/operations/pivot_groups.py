"""Grouping dates into years, quarters, months and days, and numbers into ranges.

A group is a list of items and the item each record falls in. Items start with one for
everything below the range and end with one for everything above it, as Excel's do.
"""

import calendar
import datetime as dt
import math
from dataclasses import dataclass
from typing import Literal

from excel_mcp.errors import InvalidArgumentError
from excel_mcp.operations.pivot_options import DateUnit, NumberGroup

UNIT_ORDER: tuple[DateUnit, ...] = ("years", "quarters", "months", "days")
_MONTHS = [calendar.month_abbr[month] for month in range(1, 13)]
_LEAP_YEAR = 2000
_DAYS = [dt.date(_LEAP_YEAR, 1, 1) + dt.timedelta(days=offset) for offset in range(366)]


@dataclass(frozen=True)
class Grouping:
    labels: list[str]
    """Every item, including the one below the range and the one above it."""
    positions: list[int]
    """For each record, its item's position in ``labels``."""
    by: Literal["years", "quarters", "months", "days", "range"]
    start: dt.datetime | float
    end: dt.datetime | float
    interval: float = 1
    explicit: tuple[bool, bool] = (False, False)
    """Whether start and end were given, not taken from the data."""


def group_dates(values: list[dt.datetime], units: list[DateUnit]) -> list[Grouping]:
    """One grouping per unit, largest unit first."""
    first = min(values).replace(hour=0, minute=0, second=0, microsecond=0)
    last = max(values).replace(hour=0, minute=0, second=0, microsecond=0)
    below, above = f"<{first:%d-%m-%y}", f">{last:%d-%m-%y}"
    groupings = []
    for unit in sorted(set(units), key=UNIT_ORDER.index):
        labels, positions = _date_items(unit, values, first.year)
        groupings.append(
            Grouping([below, *labels, above], [p + 1 for p in positions], unit, first, last)
        )
    return groupings


def _date_items(
    unit: DateUnit, values: list[dt.datetime], first_year: int
) -> tuple[list[str], list[int]]:
    match unit:
        case "years":
            last_year = max(value.year for value in values)
            labels = [str(year) for year in range(first_year, last_year + 1)]
            return labels, [value.year - first_year for value in values]
        case "quarters":
            return [f"Qtr{quarter}" for quarter in range(1, 5)], [
                (value.month - 1) // 3 for value in values
            ]
        case "months":
            return _MONTHS, [value.month - 1 for value in values]
        case _:
            labels = [f"{day:%d-%b}" for day in _DAYS]
            slots = {(day.month, day.day): index for index, day in enumerate(_DAYS)}
            return labels, [slots[(value.month, value.day)] for value in values]


def group_numbers(values: list[float], group: NumberGroup) -> Grouping:
    start = min(values) if group.start is None else group.start
    end = max(values) if group.end is None else group.end
    if end <= start:
        raise InvalidArgumentError(
            f"The group's end ({end:g}) must be above its start ({start:g})."
        )
    count = math.ceil((end - start) / group.by)
    if count > MAX_GROUPS:
        raise InvalidArgumentError(
            f"Grouping by {group.by:g} would make {count:,} ranges; the limit is {MAX_GROUPS:,}."
        )
    whole = float(start).is_integer() and float(group.by).is_integer()
    closed = math.isclose(start + count * group.by, end)
    labels = [f"<{_number(start)}"]
    for index in range(count):
        low, high = start + index * group.by, start + (index + 1) * group.by
        last = index == count - 1
        top = high if (last and closed) or not whole else high - 1
        labels.append(f"{_number(low)}-{_number(top)}")
    labels.append(f">{_number(start + count * group.by)}")
    positions = [_range_position(value, start, end, group.by, count) for value in values]
    return Grouping(
        labels,
        positions,
        "range",
        start,
        end,
        group.by,
        (group.start is not None, group.end is not None),
    )


MAX_GROUPS = 1_000


def _range_position(value: float, start: float, end: float, by: float, count: int) -> int:
    if value < start:
        return 0
    if value > end:
        return count + 1
    return 1 + min(int((value - start) // by), count - 1)


def _number(value: float) -> str:
    return str(int(value)) if float(value).is_integer() else repr(value)
