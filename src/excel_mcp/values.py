"""Conversion between JSON cell values and openpyxl cell values.

Dates travel as ISO 8601 strings in both directions: reading returns
``"2026-01-31"`` or ``"2026-01-31T09:30:00"``, and writing a string in exactly
that form stores a real Excel date.
"""

import datetime as dt
import re
from collections.abc import Iterable
from decimal import Decimal

from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE

from excel_mcp.errors import InvalidArgumentError
from excel_mcp.formulas import check_formula
from excel_mcp.xlfn import add_prefixes

CellValue = str | int | float | bool | None

_ISO_DATE = re.compile(r"\d{4}-\d{2}-\d{2}")
_ISO_DATETIME = re.compile(r"\d{4}-\d{2}-\d{2}T\d{2}:\d{2}(:\d{2}(\.\d+)?)?")

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
        check_formula(value, sheet_names)
        return add_prefixes(value)
    try:
        if _ISO_DATE.fullmatch(value):
            return dt.date.fromisoformat(value)
        if _ISO_DATETIME.fullmatch(value):
            return dt.datetime.fromisoformat(value)
    except ValueError:
        raise InvalidArgumentError(f"{value!r} looks like a date but is not valid.") from None
    return value


def date_number_format(value: object) -> str | None:
    """The number format to apply so a written date displays as a date."""
    if isinstance(value, dt.datetime):
        return DATETIME_FORMAT
    if isinstance(value, dt.date):
        return DATE_FORMAT
    return None
