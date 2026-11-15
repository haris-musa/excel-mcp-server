import datetime as dt

import pytest

from excel_mcp.errors import InvalidArgumentError, UnsafeFormulaError
from excel_mcp.operations.csv_values import typed_field
from excel_mcp.refs import CellRange, parse_cell, parse_range
from excel_mcp.values import parse_iso_date, to_cell, to_json, typed_value


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
    assert to_cell("2026-01-31", ["Sheet1"]) == dt.date(2026, 1, 31)
    assert to_cell("2026-01-31T09:30:00", ["Sheet1"]) == dt.datetime(2026, 1, 31, 9, 30)
    assert to_json(dt.datetime(2026, 1, 31, 9, 30)) == "2026-01-31T09:30:00"


def test_invalid_date_is_an_argument_error() -> None:
    with pytest.raises(InvalidArgumentError, match="date"):
        to_cell("2026-13-45", ["Sheet1"])


ARABIC_INDIC = str.maketrans("0123456789", "".join(chr(0x0660 + i) for i in range(10)))
FULLWIDTH = str.maketrans("0123456789", "".join(chr(0xFF10 + i) for i in range(10)))
THAI = str.maketrans("0123456789", "".join(chr(0x0E50 + i) for i in range(10)))


@pytest.mark.parametrize(
    "text",
    [
        "2026-01-31".translate(ARABIC_INDIC),
        "2026-01-31".translate(FULLWIDTH),
        "2026-01-31T10:30".translate(ARABIC_INDIC),
        "2026-" + "01".translate(ARABIC_INDIC) + "-31",
    ],
)
def test_dates_in_other_digits_stay_text(text: str) -> None:
    assert to_cell(text, ["Sheet1"]) == text
    assert parse_iso_date(text) is None


@pytest.mark.parametrize(
    ("text", "expected"),
    [
        ("123".translate(ARABIC_INDIC), (123, None)),
        ("1,234.5".translate(ARABIC_INDIC), (1234.5, "#,##0.0")),
        ("50%".translate(ARABIC_INDIC), (0.5, "0%")),
        ("2026-01-31".translate(ARABIC_INDIC), (dt.date(2026, 1, 31), "yyyy-mm-dd")),
        ("1" + "5".translate(THAI), (15, None)),
        ("123".translate(FULLWIDTH), ("123".translate(FULLWIDTH), None)),
        ("50%".translate(FULLWIDTH), ("50%".translate(FULLWIDTH), None)),
        ("2026-01-31".translate(FULLWIDTH), ("2026-01-31".translate(FULLWIDTH), None)),
        (" x" + "1".translate(ARABIC_INDIC), (" x" + "1".translate(ARABIC_INDIC), None)),
    ],
)
def test_typing_reads_the_digits_excel_reads(text: str, expected: tuple[object, object]) -> None:
    assert typed_value(text, ["Sheet1"]) == expected


@pytest.mark.parametrize(
    ("text", "expected"),
    [
        ("123".translate(ARABIC_INDIC), (123, None)),
        ("2026-01-31".translate(ARABIC_INDIC), (dt.date(2026, 1, 31), "yyyy-mm-dd")),
        ("123".translate(FULLWIDTH), ("123".translate(FULLWIDTH), None)),
    ],
)
def test_csv_fields_read_the_digits_excel_reads(text: str, expected: tuple[object, object]) -> None:
    assert typed_field(text, ".", "ymd") == expected


def test_references_read_the_digits_excel_reads() -> None:
    assert parse_cell("B" + "12".translate(ARABIC_INDIC)) == (12, 2)
    assert parse_range("A1:C" + "3".translate(THAI)) == CellRange(1, 1, 3, 3)
    with pytest.raises(InvalidArgumentError, match="Invalid range"):
        parse_cell("B" + "12".translate(FULLWIDTH))


def test_plain_values_pass_through() -> None:
    for value in ["text", "00123", 42, 1.5, True, None]:
        assert to_cell(value, ["Sheet1"]) == value


def test_formulas_are_checked() -> None:
    assert to_cell("=SUM(A1:A2)", ["Sheet1"]) == "=SUM(A1:A2)"
    with pytest.raises(UnsafeFormulaError):
        to_cell('=WEBSERVICE("https://x")', ["Sheet1"])
