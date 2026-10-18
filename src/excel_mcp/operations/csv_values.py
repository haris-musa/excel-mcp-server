"""What a CSV field becomes in a cell.

Fields are never formulas. A number, TRUE/FALSE or a date becomes a value; everything else,
including text that starts with '=', '+' or '@', stays text.
"""

import datetime as dt
import math
import re
from functools import cache
from typing import Literal

from excel_mcp.values import CellValue, date_number_format

DateOrder = Literal["ymd", "dmy", "mdy"]

_ISO_DATE = re.compile(r"\d{4}-\d{2}-\d{2}")
_ISO_DATETIME = re.compile(r"\d{4}-\d{2}-\d{2}[T ]\d{2}:\d{2}(:\d{2}(\.\d+)?)?")
_NUMERIC_DATE = re.compile(r"(\d{1,2})([/.-])(\d{1,2})\2(\d{4})")
_LEADING_ZERO = re.compile(r"-?0\d")
_EXCEL_PRECISION = 15

Typed = tuple[CellValue | dt.date | dt.datetime, str | None]


def typed_field(text: str, decimal: str, date_order: DateOrder) -> Typed:
    """The value and number format for ``text``; other text is returned as it is."""
    stripped = text.strip()
    if stripped.upper() in ("TRUE", "FALSE"):
        return stripped.upper() == "TRUE", None
    if (number := _number(stripped, decimal)) is not None:
        return number
    if (moment := _moment(stripped, date_order)) is not None:
        return moment, date_number_format(moment)
    return text, None


@cache
def _patterns(decimal: str) -> tuple[re.Pattern[str], re.Pattern[str], re.Pattern[str]]:
    point = re.escape(decimal)
    group = re.escape("," if decimal == "." else ".")
    unsigned = rf"(\d+({point}\d*)?|{point}\d+)"
    return (
        re.compile(rf"-?{unsigned}([eE][+-]?\d+)?"),
        re.compile(rf"-?\d{{1,3}}({group}\d{{3}})+({point}\d+)?"),
        re.compile(rf"-?{unsigned}%"),
    )


def _number(text: str, decimal: str) -> Typed | None:
    """A number as Excel reads it, except that text of digits Excel would alter stays text.

    That is leading zeros ('007') and more than 15 digits, which Excel rounds.
    """
    if _LEADING_ZERO.match(text) or sum(c.isdigit() for c in re.split("[eE]", text)[0]) > 15:
        return None
    plain, grouped, percent = _patterns(decimal)
    group = "," if decimal == "." else "."
    if plain.fullmatch(text):
        number, number_format = _parse(text, decimal), None
    elif grouped.fullmatch(text):
        number = _parse(text.replace(group, ""), decimal)
        number_format = "#,##0" + _places(text, decimal)
    elif percent.fullmatch(text):
        number = _parse(text[:-1], decimal) / 100
        number_format = "0" + _places(text[:-1], decimal) + "%"
    else:
        return None
    return (number, number_format) if math.isfinite(number) else None


def _parse(text: str, decimal: str) -> int | float:
    canonical = text.replace(decimal, ".")
    return float(canonical) if re.search(r"[.eE]", canonical) else int(canonical)


def _places(text: str, decimal: str) -> str:
    fraction = text.partition(decimal)[2]
    return "." + "0" * len(fraction) if fraction else ""


def _moment(text: str, date_order: DateOrder) -> dt.date | dt.datetime | None:
    """A date as ISO 8601, or numeric 31/12/2026 only in the order the caller names."""
    try:
        if _ISO_DATE.fullmatch(text):
            moment = dt.date.fromisoformat(text)
        elif _ISO_DATETIME.fullmatch(text):
            moment = dt.datetime.fromisoformat(text)
        elif date_order != "ymd" and (parts := _NUMERIC_DATE.fullmatch(text)):
            first, second = int(parts[1]), int(parts[3])
            day, month = (first, second) if date_order == "dmy" else (second, first)
            moment = dt.date(int(parts[4]), month, day)
        else:
            return None
    except ValueError:
        return None
    return moment if moment.year >= 1900 else None
