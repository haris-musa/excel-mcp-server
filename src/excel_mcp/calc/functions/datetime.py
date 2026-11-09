"""Date and time functions. Serial numbers are Excel's 1900 date system."""

import calendar
import datetime as dt

from excel_mcp.calc.dateparts import LAST_SERIAL, month_day_year, serial_day
from excel_mcp.calc.functions.helpers import flat
from excel_mcp.calc.registry import function
from excel_mcp.calc.values import (
    NUM,
    VALUE,
    ExcelError,
    FormulaError,
    Scalar,
    UncalculableError,
    Value,
    date_to_serial,
    is_number,
    parse_number,
    serial_to_date,
    to_int,
    to_number,
    to_text,
)

_MAX_STEPS = 1_000_000


def _date(value: Scalar) -> dt.date:
    return serial_to_date(to_number(value))


def _serial(date: dt.date) -> float:
    if date.year > 9999:
        raise FormulaError(NUM)
    return float(date_to_serial(date))


def add_months(date: dt.date, months: int) -> dt.date:
    year, month = divmod(date.year * 12 + date.month - 1 + months, 12)
    if not 1900 <= year <= 9999:
        raise FormulaError(NUM)
    day = min(date.day, calendar.monthrange(year, month + 1)[1])
    return dt.date(year, month + 1, day)


@function("DATE", kind="scalar")
def date_(year: Scalar, month: Scalar, day: Scalar) -> float:
    y, m, d = to_int(year), to_int(month), to_int(day)
    if 0 <= y <= 1899:
        y += 1900
    if not 1900 <= y <= 9999:
        raise FormulaError(NUM)
    y, m0 = divmod(y * 12 + m - 1, 12)
    if not 1 <= y <= 9999:
        raise FormulaError(NUM)
    result = dt.date(y, m0 + 1, 1) + dt.timedelta(days=d - 1)
    serial = float(date_to_serial(result)) if result.year >= 1900 else -1.0
    if serial < 0:
        raise FormulaError(NUM)
    return serial


def _ymd(serial: Scalar) -> tuple[int, int, int]:
    days = serial_day(to_number(serial))
    if days < 0 or days > LAST_SERIAL:
        raise FormulaError(NUM)
    return month_day_year(days)


@function("YEAR", kind="scalar")
def year(serial: Scalar) -> float:
    return float(_ymd(serial)[0])


@function("MONTH", kind="scalar")
def month(serial: Scalar) -> float:
    return float(_ymd(serial)[1])


@function("DAY", kind="scalar")
def day(serial: Scalar) -> float:
    return float(_ymd(serial)[2])


def _seconds(serial: Scalar) -> int:
    number = to_number(serial)
    if number < 0 or number >= LAST_SERIAL + 1:
        raise FormulaError(NUM)
    return round((number % 1) * 86400) % 86400


@function("HOUR", kind="scalar")
def hour(serial: Scalar) -> float:
    return float(_seconds(serial) // 3600)


@function("MINUTE", kind="scalar")
def minute(serial: Scalar) -> float:
    return float(_seconds(serial) // 60 % 60)


@function("SECOND", kind="scalar")
def second(serial: Scalar) -> float:
    return float(_seconds(serial) % 60)


@function("TIME", kind="scalar")
def time_(hours: Scalar, minutes: Scalar, seconds: Scalar) -> float:
    h, m, s = to_int(hours), to_int(minutes), to_int(seconds)
    total = h * 3600 + m * 60 + s
    if h > 32767 or m > 32767 or s > 32767 or total < 0:
        raise FormulaError(NUM)
    return (total % 86400) / 86400


@function("TODAY")
def today() -> float:
    return _serial(dt.date.today())


@function("NOW")
def now() -> float:
    moment = dt.datetime.now()
    return (
        _serial(moment.date())
        + (moment - dt.datetime.combine(moment.date(), dt.time())).total_seconds() / 86400
    )


@function("EDATE", kind="scalar")
def edate(start: Scalar, months: Scalar) -> float:
    return _serial(add_months(_date(start), to_int(months)))


@function("EOMONTH", kind="scalar")
def eomonth(start: Scalar, months: Scalar) -> float:
    shifted = add_months(_date(start), to_int(months))
    return _serial(shifted.replace(day=calendar.monthrange(shifted.year, shifted.month)[1]))


@function("DATEVALUE", kind="scalar")
def datevalue(text: Scalar) -> float:
    number = parse_number(to_text(text))
    if number is None:
        raise FormulaError(VALUE)
    return float(int(number))


@function("TIMEVALUE", kind="scalar")
def timevalue(text: Scalar) -> float:
    number = parse_number(to_text(text))
    if number is None:
        raise FormulaError(VALUE)
    return number % 1


@function("DAYS", kind="scalar")
def days(end: Scalar, start: Scalar) -> float:
    first, last = int(to_number(start)), int(to_number(end))
    if not 0 <= first <= LAST_SERIAL or not 0 <= last <= LAST_SERIAL:
        raise FormulaError(NUM)
    return float(last - first)


def days360(start: dt.date, end: dt.date, european: bool) -> int:
    first, last = start.day, end.day
    if not european:
        if _last_of_february(end) and _last_of_february(start):
            last = 30
        if _last_of_february(start):
            first = 30
        if last == 31 and first >= 30:
            last = 30
    elif last == 31:
        last = 30
    if first == 31:
        first = 30
    return (end.year - start.year) * 360 + (end.month - start.month) * 30 + last - first


def _last_of_february(date: dt.date) -> bool:
    return date.month == 2 and date.day == calendar.monthrange(date.year, 2)[1]


@function("DAYS360", kind="scalar")
def days360_(start: Scalar, end: Scalar, method: Scalar = False) -> float:
    european = bool(to_number(method))
    return float(days360(_date(start), _date(end), european))


def _leap(year_: int) -> bool:
    return calendar.isleap(year_)


@function("YEARFRAC", kind="scalar")
def yearfrac(start: Scalar, end: Scalar, basis: Scalar = 0) -> float:
    return year_fraction(_date(start), _date(end), to_int(basis))


def year_fraction(start: dt.date, end: dt.date, basis: int) -> float:
    if start > end:
        start, end = end, start
    span = (end - start).days
    match basis:
        case 0:
            return days360(start, end, european=False) / 360
        case 1:
            return span / actual_year_length(start, end)
        case 2:
            return span / 360
        case 3:
            return span / 365
        case 4:
            return days360(start, end, european=True) / 360
    raise FormulaError(NUM)


def actual_year_length(start: dt.date, end: dt.date) -> float:
    if start.year == end.year:
        return 366 if _leap(start.year) else 365
    anniversary = (
        start.replace(year=start.year + 1)
        if not (start.month == 2 and start.day == 29)
        else dt.date(start.year + 1, 3, 1)
    )
    if end <= anniversary:
        february = (_leap(start.year) and start <= dt.date(start.year, 2, 29)) or (
            _leap(end.year) and end >= dt.date(end.year, 2, 29)
        )
        return 366 if february else 365
    years = range(start.year, end.year + 1)
    return sum(366 if _leap(y) else 365 for y in years) / len(years)


@function("DATEDIF", kind="scalar")
def datedif(start: Scalar, end: Scalar, unit: Scalar) -> float:
    first, last = _date(start), _date(end)
    if first > last:
        raise FormulaError(NUM)
    months = (last.year - first.year) * 12 + last.month - first.month - (last.day < first.day)
    match to_text(unit).upper():
        case "Y":
            return float(months // 12)
        case "M":
            return float(months)
        case "D":
            return float((last - first).days)
        case "YM":
            return float(months % 12)
        case "MD":
            if last.day >= first.day:
                return float(last.day - first.day)
            previous = add_months(last.replace(day=1), -1)
            return float(
                last.day + calendar.monthrange(previous.year, previous.month)[1] - first.day
            )
        case "YD":
            return float(_days_in_year_cycle(first, last))
    raise FormulaError(NUM)


def _days_in_year_cycle(first: dt.date, last: dt.date) -> int:
    """DATEDIF "YD": Excel puts the end's month and day into the first year it can be reached
    from, except from 29 February to 30 March of a leap year, where it counts from the last
    anniversary."""
    if calendar.isleap(last.year) and (2, 29) <= (last.month, last.day) <= (3, 30):
        year = last.year - ((last.month, last.day) < (first.month, first.day))
        day = min(first.day, calendar.monthrange(year, first.month)[1])
        return (last - dt.date(year, first.month, day)).days
    year = first.year + ((last.month, last.day) < (first.month, first.day))
    day = min(last.day, calendar.monthrange(year, last.month)[1])
    return (dt.date(year, last.month, day) - first).days


_WEEK_START = {1: 6, 2: 0, **{n: n - 11 for n in range(11, 18)}}


@function("WEEKDAY", kind="scalar")
def weekday(serial: Scalar, kind: Scalar = 1) -> float:
    days = serial_day(to_number(serial))
    if days < 0 or days > LAST_SERIAL:
        raise FormulaError(NUM)
    day_index = (days - 2) % 7  # Excel's weekdays are those of its own calendar
    match to_int(kind):
        case 1:
            return float((day_index + 1) % 7 + 1)
        case 2:
            return float(day_index + 1)
        case 3:
            return float(day_index)
        case n if 11 <= n <= 17:
            return float((day_index - (n - 11)) % 7 + 1)
    raise FormulaError(NUM)


@function("WEEKNUM", kind="scalar")
def weeknum(serial: Scalar, kind: Scalar = 1) -> float:
    date = _date(serial)
    if date.year == 1900:
        raise UncalculableError("week numbers in 1900, whose calendar Excel shifts")
    system = to_int(kind)
    if system == 21:
        return float(date.isocalendar().week)
    if system not in _WEEK_START:
        raise FormulaError(NUM)
    first = dt.date(date.year, 1, 1)
    offset = (first.weekday() - _WEEK_START[system]) % 7
    return float((date.timetuple().tm_yday - 1 + offset) // 7 + 1)


@function("ISOWEEKNUM", kind="scalar")
def isoweeknum(serial: Scalar) -> float:
    date = _date(serial)
    if date.year == 1900:
        raise UncalculableError("week numbers in 1900, whose calendar Excel shifts")
    return float(date.isocalendar().week)


_WEEKENDS = {
    1: "0000011", 2: "1000001", 3: "1100000", 4: "0110000", 5: "0011000", 6: "0001100",
    7: "0000110", 11: "0000001", 12: "1000000", 13: "0100000", 14: "0010000",
    15: "0001000", 16: "0000100", 17: "0000010",
}  # fmt: skip


def _weekend_mask(weekend: Scalar) -> str:
    if isinstance(weekend, str):
        if len(weekend) != 7 or set(weekend) - {"0", "1"}:
            raise FormulaError(VALUE)
        return weekend
    code = to_int(weekend)
    if code not in _WEEKENDS:
        raise FormulaError(NUM)
    return _WEEKENDS[code]


def _holidays(value: Value) -> set[int]:
    found: set[int] = set()
    for item in flat([value]):
        if isinstance(item, ExcelError):
            raise FormulaError(item)
        if item is None:
            continue
        if not is_number(item):
            raise FormulaError(VALUE)
        found.add(int(item))  # pyright: ignore[reportArgumentType]
    return found


def _working(serial: int, mask: str, holidays: set[int]) -> bool:
    return mask[serial_to_date(serial).weekday()] == "0" and serial not in holidays


def _networkdays(start: Scalar, end: Scalar, mask: str, holidays: Value) -> float:
    first, last = int(to_number(start)), int(to_number(end))
    skipped = _holidays(holidays)
    sign = 1
    if first > last:
        first, last, sign = last, first, -1
    if last - first > _MAX_STEPS:
        raise UncalculableError("date range too long")
    return float(sign * sum(_working(n, mask, skipped) for n in range(first, last + 1)))


def _workday(start: Scalar, count: Scalar, mask: str, holidays: Value) -> float:
    current, remaining = int(to_number(start)), to_int(count)
    skipped = _holidays(holidays)
    step = 1 if remaining >= 0 else -1
    if abs(remaining) > _MAX_STEPS:
        raise UncalculableError("date range too long")
    while remaining:
        current += step
        if _working(current, mask, skipped):
            remaining -= step
    serial_to_date(current)
    return float(current)


@function("NETWORKDAYS")
def networkdays(start: Scalar, end: Scalar, holidays: Value = None) -> float:
    return _networkdays(start, end, _WEEKENDS[1], holidays)


@function("NETWORKDAYS.INTL")
def networkdays_intl(
    start: Scalar, end: Scalar, weekend: Scalar = 1, holidays: Value = None
) -> float:
    return _networkdays(start, end, _weekend_mask(weekend), holidays)


@function("WORKDAY")
def workday(start: Scalar, count: Scalar, holidays: Value = None) -> float:
    return _workday(start, count, _WEEKENDS[1], holidays)


@function("WORKDAY.INTL")
def workday_intl(
    start: Scalar, count: Scalar, weekend: Scalar = 1, holidays: Value = None
) -> float:
    mask = _weekend_mask(weekend)
    if mask == "1111111":
        raise FormulaError(VALUE)
    return _workday(start, count, mask, holidays)
