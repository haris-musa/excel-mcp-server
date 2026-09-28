import datetime as dt

import pytest

from excel_mcp.errors import InvalidArgumentError, UnsafeFormulaError
from excel_mcp.refs import CellRange, parse_cell, parse_range
from excel_mcp.values import to_cell, to_json


@pytest.mark.parametrize(
    ("ref", "expected"),
    [
        ("B2", CellRange(2, 2, 2, 2)),
        ("a1:c3", CellRange(1, 1, 3, 3)),
        ("$A$1:$B$2", CellRange(1, 1, 2, 2)),
        ("C3:A1", CellRange(1, 1, 3, 3)),
        ("AA10", CellRange(10, 27, 10, 27)),
    ],
)
def test_parse_range(ref: str, expected: CellRange) -> None:
    assert parse_range(ref) == expected


@pytest.mark.parametrize("ref", ["", "A", "1:3", "A:A", "A0", "ZZZZ1", "Sheet1!A1", "A1:B"])
def test_parse_range_rejects_invalid_or_unbounded(ref: str) -> None:
    with pytest.raises(InvalidArgumentError):
        parse_range(ref)


def test_parse_cell_rejects_ranges() -> None:
    assert parse_cell("C7") == (7, 3)
    with pytest.raises(InvalidArgumentError, match="single cell"):
        parse_cell("A1:B2")


def test_cell_range_str() -> None:
    assert str(CellRange(1, 1, 1, 1)) == "A1"
    assert str(CellRange(2, 1, 10, 4)) == "A2:D10"


def test_dates_round_trip_as_iso_strings() -> None:
    assert to_cell("2026-01-31") == dt.date(2026, 1, 31)
    assert to_cell("2026-01-31T09:30:00") == dt.datetime(2026, 1, 31, 9, 30)
    assert to_json(dt.datetime(2026, 1, 31, 9, 30)) == "2026-01-31T09:30:00"


def test_invalid_date_is_an_argument_error() -> None:
    with pytest.raises(InvalidArgumentError, match="date"):
        to_cell("2026-13-45")


def test_plain_values_pass_through() -> None:
    for value in ["text", "00123", 42, 1.5, True, None]:
        assert to_cell(value) == value


def test_formulas_are_checked() -> None:
    assert to_cell("=SUM(A1:A2)") == "=SUM(A1:A2)"
    with pytest.raises(UnsafeFormulaError):
        to_cell('=WEBSERVICE("https://x")')
