"""The value model of the calculator and Excel's coercion rules."""

import datetime as dt
import re
import sys
from dataclasses import dataclass
from decimal import ROUND_DOWN, Context, Decimal
from typing import Any

from excel_mcp.calc.dateparts import LAST_SERIAL, serial_day
from excel_mcp.refs import MAX_COLUMN, MAX_ROW

DATE_ORIGIN = dt.date(1899, 12, 30)
_MAX_PLAIN_LENGTH = 20
_FIRST_15_DIGITS = Context(prec=15, rounding=ROUND_DOWN)


@dataclass(frozen=True)
class ExcelError:
    code: str

    def __str__(self) -> str:
        return self.code


DIV0 = ExcelError("#DIV/0!")
NA = ExcelError("#N/A")
VALUE = ExcelError("#VALUE!")
REF = ExcelError("#REF!")
NAME = ExcelError("#NAME?")
NUM = ExcelError("#NUM!")
NULL = ExcelError("#NULL!")
CALC = ExcelError("#CALC!")
ERRORS = {e.code: e for e in (DIV0, NA, VALUE, REF, NAME, NUM, NULL, CALC)}
ERROR_NUMBERS = {
    "#NULL!": 1,
    "#DIV/0!": 2,
    "#VALUE!": 3,
    "#REF!": 4,
    "#NAME?": 5,
    "#NUM!": 6,
    "#N/A": 7,
}

Scalar = float | int | str | bool | None | ExcelError


class Grid:
    """A rectangle of values: a range reference or an array result."""

    def __init__(self, rows: list[list[Any]]) -> None:
        self.rows = rows

    @property
    def height(self) -> int:
        return len(self.rows)

    @property
    def width(self) -> int:
        return len(self.rows[0]) if self.rows else 0

    def flat(self) -> list["Value"]:
        return [value for row in self.rows for value in row]

    def columns(self) -> list[list["Value"]]:
        return [list(column) for column in zip(*self.rows, strict=True)]


class RefGrid(Grid):
    """A grid read from the sheet, which remembers where it came from."""

    def __init__(
        self,
        rows: list[list["Value"]],
        top: int,
        left: int,
        sheets: int,
        clipped_rows: bool = False,
        clipped_cols: bool = False,
        sheet: str | None = None,
    ) -> None:
        super().__init__(rows)
        self.sheet = sheet  # the title of the one sheet it was read from
        self.clipped_rows = clipped_rows
        self.clipped_cols = clipped_cols
        self.top = top
        self.left = left
        self.sheets = sheets

    @property
    def total_rows(self) -> int:
        """The rows of the reference itself: a whole column has them all, used or not."""
        return MAX_ROW if self.clipped_rows else self.height

    @property
    def total_cols(self) -> int:
        return MAX_COLUMN if self.clipped_cols else self.width


Value = Scalar | Grid


class FormulaError(Exception):
    """Evaluation produced an Excel error value."""

    def __init__(self, error: ExcelError) -> None:
        super().__init__(error.code)
        self.error = error


class UncalculableError(Exception):
    """The calculator cannot reproduce Excel's result, so the cell is left out."""


MAX_TEXT = 32_767  # characters in a cell, as in Excel
MAX_ARRAY_CELLS = 100_000
MAX_PERIODS = 10_000  # iterations of the depreciation and cumulative loan functions


def check_periods(count: float) -> None:
    if count > MAX_PERIODS:
        raise UncalculableError("too many periods")


def check_array_size(height: int, width: int) -> None:
    """Refuse an array result before it is built."""
    if height * width > MAX_ARRAY_CELLS:
        raise UncalculableError("array too large")


def number_literal(text: str) -> float:
    """A number typed in a formula or coerced from text: Excel keeps its first 15 digits."""
    return flush(float(_FIRST_15_DIGITS.create_decimal(text)))


def plain_number(text: str) -> float | None:
    """The number written in text such as ``-1.5E+3``; None where Excel reads no number.

    Excel keeps the first 15 digits, and takes no text beyond these magnitudes for a number.
    """
    digits = _FIRST_15_DIGITS.create_decimal(text)
    number = float(digits)
    in_range = 1e-309 <= abs(number) < 1e308
    if not in_range and (digits != 0 or not -309 <= digits.as_tuple().exponent <= 307):  # pyright: ignore[reportOperatorIssue]
        return None
    return flush(number)


def flush(number: float) -> float:
    """Excel's multiplication and division give 0 for a result below the smallest normal
    double, and never -0. Addition, subtraction and most functions keep such results."""
    return 0.0 if abs(number) < sys.float_info.min else number


def is_number(value: object) -> bool:
    return isinstance(value, int | float) and not isinstance(value, bool)


def scalar(value: Value) -> Scalar:
    """The value itself, or the only cell of a 1x1 grid."""
    if isinstance(value, Grid):
        if value.height != 1 or value.width != 1:
            raise UncalculableError("range where a single value is expected")
        return value.rows[0][0]  # pyright: ignore[reportReturnType]
    return value


def serial_to_date(serial: float) -> dt.date:
    days = serial_day(serial)
    if days < 0 or days > LAST_SERIAL:
        raise FormulaError(NUM)
    if days < 61:
        # Excel's 1900 leap-year bug: serial 0 is "January 0" and 60 the non-existent 29 February.
        raise UncalculableError("dates before 1900-03-01")
    return DATE_ORIGIN + dt.timedelta(days=days)


def date_to_serial(date: dt.date) -> int:
    if date < dt.date(1900, 3, 1):
        raise UncalculableError("dates before 1900-03-01")
    return (date - DATE_ORIGIN).days


_PLAIN_NUMBER = re.compile(r"(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?")
_THOUSANDS = re.compile(r"\d{1,3}(,\d{3})+(\.\d*)?")
_ISO_DATE = re.compile(r"(\d{4})-(\d{1,2})-(\d{1,2})")
_TIME = re.compile(r"(\d{1,2}):(\d{2})(?::(\d{2}))?")


_DATE_TEXT = re.compile(
    r"\d\s*[/-]\s*\d|\d\s*[AaPp][Mm]|(?i:jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\.?\s*\d|\d\s*(?i:jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)"
)


def parse_number(text: str) -> float | None:
    """Excel's conversion of text to a number, or None when it is not one."""
    body = text.strip()
    if not body:
        return None
    percent = body.endswith("%")
    if percent:
        body = body[:-1].rstrip()
    sign = ""
    if body[:1] in ("+", "-"):
        sign, body = body[0], body[1:].lstrip()
    body = body.removeprefix("$")
    if _THOUSANDS.fullmatch(body):
        body = body.replace(",", "")
    if _PLAIN_NUMBER.fullmatch(body):
        number = plain_number(sign + body)
        return None if number is None else number / 100 if percent else number
    if not percent and not sign and (moment := _parse_date_time(body)) is not None:
        return moment
    if _DATE_TEXT.search(text):
        raise UncalculableError("text that Excel may read as a date")
    return None


def _parse_date_time(text: str) -> float | None:
    day_part, _, time_part = text.partition(" ")
    serial = 0.0
    if time_part or ":" in text:
        match = _TIME.fullmatch(time_part or day_part)
        if not match:
            return None
        hours, minutes, seconds = int(match[1]), int(match[2]), int(match[3] or 0)
        if hours > 23 or minutes > 59 or seconds > 59:
            return None
        serial = (hours * 3600 + minutes * 60 + seconds) / 86400
        if not time_part:
            return serial
    match = _ISO_DATE.fullmatch(day_part)
    if not match:
        return None
    try:
        date = dt.date(int(match[1]), int(match[2]), int(match[3]))
    except ValueError:
        return None
    return date_to_serial(date) + serial


def to_number(value: Value) -> float:
    value = scalar(value)
    if isinstance(value, bool):
        return float(value)
    if isinstance(value, int | float):
        return value
    if value is None:
        return 0.0
    if isinstance(value, ExcelError):
        raise FormulaError(value)
    number = parse_number(value)
    if number is None:
        raise FormulaError(VALUE)
    return number


def to_int(value: Value) -> int:
    """Truncate toward zero like Excel does for integer arguments."""
    return int(to_number(value))


def to_bool(value: Value) -> bool:
    value = scalar(value)
    if isinstance(value, str):
        match value.upper():
            case "TRUE":
                return True
            case "FALSE":
                return False
        raise FormulaError(VALUE)
    if isinstance(value, ExcelError):
        raise FormulaError(value)
    return bool(value)


def number_text(number: float) -> str:
    """Excel's General text for a number: 15 significant digits, in E notation when long."""
    if abs(number) < sys.float_info.min:
        return "0"
    digits = Decimal(f"{number:.15g}")
    fixed = f"{abs(digits):f}"
    if "." in fixed:
        fixed = fixed.rstrip("0").rstrip(".")
    if len(fixed) <= _MAX_PLAIN_LENGTH:
        return ("-" if number < 0 else "") + fixed
    mantissa, _, exponent = f"{abs(digits):.14E}".partition("E")
    mantissa = mantissa.rstrip("0").rstrip(".")
    sign = "-" if number < 0 else ""
    return f"{sign}{mantissa}E{exponent[0]}{exponent[1:].lstrip('0').rjust(2, '0')}"


def to_text(value: Value) -> str:
    value = scalar(value)
    if value is None:
        return ""
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, int | float):
        return number_text(value)
    if isinstance(value, ExcelError):
        raise FormulaError(value)
    return value


def type_rank(value: Scalar) -> int:
    if isinstance(value, bool):
        return 2
    return 1 if isinstance(value, str) else 0


def compare(left: Scalar, right: Scalar) -> int:
    """Order two values like Excel: numbers, then text (case-blind), then booleans."""
    if left is None:
        left = _blank_like(right)
    if right is None:
        right = _blank_like(left)
    if type_rank(left) != type_rank(right):
        return -1 if type_rank(left) < type_rank(right) else 1
    if isinstance(left, str) and isinstance(right, str):
        left, right = left.casefold(), right.casefold()
    elif isinstance(left, float | int) and not isinstance(left, bool):
        # Excel compares numbers at 15 significant digits: 0.1+0.2 equals 0.3.
        left, right = float(f"{left:.15g}"), float(f"{right:.15g}")
    return (left > right) - (left < right)  # pyright: ignore[reportOperatorIssue]


def _blank_like(other: Scalar) -> Scalar:
    if isinstance(other, str):
        return ""
    return False if isinstance(other, bool) else 0
