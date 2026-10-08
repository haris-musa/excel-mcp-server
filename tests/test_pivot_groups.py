"""Group items and record placement, with the labels Excel produced for the same settings."""

import datetime as dt

import pytest

from excel_mcp.errors import InvalidArgumentError
from excel_mcp.operations.pivot_groups import group_dates, group_numbers
from excel_mcp.operations.pivot_options import NumberGroup

UNITS = [10.0, 5.0, 7.0, 3.0, 12.0, 8.0, 6.0, 4.0, 9.0, 11.0]


@pytest.mark.parametrize(
    ("group", "labels"),
    [
        (NumberGroup(by=5, start=0, end=20), ["<0", "0-4", "5-9", "10-14", "15-20", ">20"]),
        (NumberGroup(by=4), ["<3", "3-6", "7-10", "11-14", ">15"]),
        (NumberGroup(by=4, start=3, end=12), ["<3", "3-6", "7-10", "11-14", ">15"]),
    ],
)
def test_number_ranges_are_labelled_like_excel(group: NumberGroup, labels: list[str]) -> None:
    assert group_numbers(UNITS, group).labels == labels


def test_fractional_ranges_and_values_outside() -> None:
    prices = [1.5 + step for step in range(10)]
    grouping = group_numbers(prices, NumberGroup(by=2.5, start=1, end=10))
    assert grouping.labels == ["<1", "1-3.5", "3.5-6", "6-8.5", "8.5-11", ">11"]
    assert grouping.positions == [1, 1, 2, 2, 2, 3, 3, 4, 4, 5]


def test_numbers_are_placed_in_their_range() -> None:
    grouping = group_numbers(UNITS, NumberGroup(by=5, start=0, end=20))
    assert grouping.positions == [3, 2, 2, 1, 3, 2, 2, 1, 2, 3]
    low = group_numbers([-3.0, 2.0, 30.0], NumberGroup(by=5, start=0, end=20))
    assert low.positions == [0, 1, 5]


def test_too_many_ranges_are_refused() -> None:
    with pytest.raises(InvalidArgumentError, match="limit"):
        group_numbers([0.0, 1_000_000.0], NumberGroup(by=1))
    with pytest.raises(InvalidArgumentError, match="above its start"):
        group_numbers([1.0, 1.0], NumberGroup(by=1))


def test_date_units_are_ordered_largest_first_and_cover_every_slot() -> None:
    values = [dt.datetime(2024, 1, 15), dt.datetime(2025, 3, 14), dt.datetime(2024, 2, 29)]
    groupings = group_dates(values, ["months", "years", "quarters"])
    assert [grouping.by for grouping in groupings] == ["years", "quarters", "months"]
    years, quarters, months = groupings
    assert years.labels[1:-1] == ["2024", "2025"]
    assert quarters.labels[1:-1] == ["Qtr1", "Qtr2", "Qtr3", "Qtr4"]
    assert months.labels[1:-1][:3] == ["Jan", "Feb", "Mar"]
    assert months.positions == [1, 3, 2]
    assert years.positions == [1, 2, 1]
    days = group_dates(values, ["days"])[0]
    assert len(days.labels) == 366 + 2
    assert days.labels[days.positions[2]] == "29-Feb"
