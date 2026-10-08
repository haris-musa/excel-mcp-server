"""Conversion between JSON cell values and openpyxl cell values.

Dates travel as ISO 8601 strings in both directions: reading returns
``"2026-01-31"`` or ``"2026-01-31T09:30:00"``, and writing a string in exactly
that form stores a real Excel date.
"""

import datetime as dt
import re
from collections.abc import Iterable
from decimal import Decimal

from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE, Cell
from openpyxl.worksheet.formula import ArrayFormula

from excel_mcp.errors import InvalidArgumentError
from excel_mcp.formulas import storable_formula

CellValue = str | int | float | bool | None

_ISO_DATE = re.compile(r"\d{4}-\d{2}-\d{2}")
_ISO_DATETIME = re.compile(r"\d{4}-\d{2}-\d{2}T\d{2}:\d{2}(:\d{2}(\.\d+)?)?")
_NUMBER = re.compile(r"[+-]?(\d+(\.\d*)?|\.\d+)([eE][+-]?\d+)?")
_THOUSANDS = re.compile(r"[+-]?\d{1,3}(,\d{3})+(\.\d+)?")
_PERCENT = re.compile(r"[+-]?(\d+(\.\d*)?|\.\d+)%")

DATE_FORMAT = "yyyy-mm-dd"
DATETIME_FORMAT = "yyyy-mm-dd hh:mm:ss"


def to_json(value: object) -> CellValue:
    """Convert a value read from openpyxl to a JSON-friendly value."""
    match value:
        case None | bool() | int() | float() | str():
            return value
        case dt.datetime() if value.time() == dt.time():
            return value.date().isoformat()
        case dt.datetime() | dt.date() | dt.time():
            return value.isoformat()
        case dt.timedelta():
            return value.total_seconds()
        case Decimal():
            return float(value)
        case ArrayFormula():
            return str(value.text)
        case _:
            return str(value)


def to_cell(value: CellValue, sheet_names: Iterable[str]) -> CellValue | dt.date | dt.datetime:
    """Convert a JSON value to what openpyxl should store.

    Strings starting with ``=`` are formulas and must pass the safety policy for
    a workbook with ``sheet_names``.
    """
    if not isinstance(value, str):
        return value
    if ILLEGAL_CHARACTERS_RE.search(value):
        raise InvalidArgumentError("Text cannot contain control characters such as '\x01'.")
    if value.startswith("="):
        return storable_formula(value, sheet_names)
    try:
        if _ISO_DATE.fullmatch(value):
            return dt.date.fromisoformat(value)
        if _ISO_DATETIME.fullmatch(value):
            return dt.datetime.fromisoformat(value)
    except ValueError:
        raise InvalidArgumentError(f"{value!r} looks like a date but is not valid.") from None
    return value


def store_value(cell: Cell, value: CellValue) -> None:
    """Store a computed value, keeping text as text even if it starts with '='."""
    cell.value = value
    if isinstance(value, str):
        cell.data_type = "s"


def date_number_format(value: object) -> str | None:
    """The number format to apply so a written date displays as a date."""
    if isinstance(value, dt.datetime):
        return DATETIME_FORMAT
    if isinstance(value, dt.date):
        return DATE_FORMAT
    return None


def parse_iso_date(text: str) -> dt.date | None:
    """The date written as ``2026-01-31``, or None if ``text`` is not in that form."""
    if not _ISO_DATE.fullmatch(text):
        return None
    try:
        return dt.date.fromisoformat(text)
    except ValueError:
        raise InvalidArgumentError(f"{text!r} looks like a date but is not valid.") from None


def typed_value(
    text: str, sheet_names: Iterable[str]
) -> tuple[CellValue | dt.date | dt.datetime, str | None]:
    """What Excel stores when ``text`` is typed into a General cell, and the format it applies.

    Numbers (also with thousands separators or a percent sign), TRUE/FALSE, ISO dates and
    formulas become values; anything else stays text.
    """
    stripped = text.strip()
    if stripped.upper() in ("TRUE", "FALSE"):
        return stripped.upper() == "TRUE", None
    if _NUMBER.fullmatch(stripped):
        return _number(stripped), None
    if _THOUSANDS.fullmatch(stripped):
        return _number(stripped.replace(",", "")), "#,##0" + _decimals(stripped)
    if _PERCENT.fullmatch(stripped):
        return _number(stripped[:-1]) / 100, "0" + _decimals(stripped[:-1]) + "%"
    value = to_cell(text, sheet_names)
    return value, date_number_format(value)


def _number(text: str) -> int | float:
    if text.lstrip("+-").isdigit() and len(text.lstrip("+-")) <= 15:
        return int(text)
    return float(text)


def _decimals(number: str) -> str:
    fraction = number.partition(".")[2]
    return "." + "0" * len(fraction) if fraction else ""
