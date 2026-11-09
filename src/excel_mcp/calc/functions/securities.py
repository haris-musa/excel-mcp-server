"""Financial functions for bonds and other securities."""

import calendar
import datetime as dt
from dataclasses import dataclass, replace

from excel_mcp.calc.functions.datetime import add_months, days360, year_fraction
from excel_mcp.calc.functions.financial import newton
from excel_mcp.calc.functions.math import naive_sum
from excel_mcp.calc.registry import function
from excel_mcp.calc.values import (
    NUM,
    VALUE,
    FormulaError,
    Scalar,
    UncalculableError,
    serial_to_date,
    to_int,
    to_number,
)


def as_date(value: Scalar) -> dt.date:
    if isinstance(value, bool):
        raise FormulaError(VALUE)
    return serial_to_date(int(to_number(value)))


def as_basis(value: Scalar) -> int:
    basis = to_int(value)
    if not 0 <= basis <= 4:
        raise FormulaError(NUM)
    return basis


def require(condition: bool) -> None:
    if not condition:
        raise FormulaError(NUM)


@dataclass(frozen=True)
class Coupons:
    """The coupon schedule of a security between settlement and maturity."""

    settlement: dt.date
    previous: dt.date
    next: dt.date
    plain_previous: dt.date
    count: int
    frequency: int
    basis: int

    @property
    def days_in_period(self) -> float:
        if self.basis == 1:
            # Excel measures the period from the unadjusted previous coupon date.
            months = 12 // self.frequency
            return float((add_months(self.plain_previous, months) - self.plain_previous).days)
        return (365 if self.basis == 3 else 360) / self.frequency

    @property
    def days_since_previous(self) -> float:
        return _between(self.previous, self.settlement, self.basis)

    @property
    def days_to_next(self) -> float:
        """What COUPDAYSNC reports."""
        if self.basis == 0:
            return self.days_remaining
        return _between(self.settlement, self.next, self.basis)

    @property
    def pricing_period(self) -> float:
        """The days in the period as pricing and duration count them: Actual/Actual measures
        between the actual coupon dates, where COUPDAYS measures from the unadjusted one."""
        if self.basis == 1:
            return float((self.next - self.previous).days)
        return self.days_in_period

    @property
    def days_remaining(self) -> float:
        """The days to the next coupon as pricing counts them."""
        if self.basis == 1:
            return float((self.next - self.settlement).days)
        return self.days_in_period - self.days_since_previous


def _between(start: dt.date, end: dt.date, basis: int) -> float:
    if basis == 0:
        return float(bond_days360(start, end))
    if basis == 4:
        return float(days360(start, end, european=True))
    return float((end - start).days)


def _last_of_february(date: dt.date) -> bool:
    return date.month == 2 and date.day == calendar.monthrange(date.year, 2)[1]


def bond_days360(start: dt.date, end: dt.date) -> int:
    """30/360 as the coupon functions count it: a month-end 31st is tested before February."""
    first, last = start.day, end.day
    if last == 31 and first >= 30:
        last = 30
    if first == 31:
        first = 30
    if _last_of_february(start):
        first = 30
        if _last_of_february(end):
            last = 30
    return (end.year - start.year) * 360 + (end.month - start.month) * 30 + last - first


def end_of_month(date: dt.date) -> dt.date:
    return date.replace(day=calendar.monthrange(date.year, date.month)[1])


def coupons(settlement: Scalar, maturity: Scalar, frequency: Scalar, basis: Scalar) -> Coupons:
    start, end = as_date(settlement), as_date(maturity)
    freq, base = to_int(frequency), as_basis(basis)
    require(freq in (1, 2, 4) and start < end)
    step = 12 // freq
    month_end = end == end_of_month(end)

    def coupon_date(periods_back: int) -> dt.date:
        date = add_months(end, -periods_back * step)
        return end_of_month(date) if month_end else date

    periods = 0
    while coupon_date(periods) > start:
        periods += 1
    plain_previous = add_months(end, -periods * step)
    return Coupons(
        start, coupon_date(periods), coupon_date(periods - 1), plain_previous, periods, freq, base
    )


@function("COUPNCD", kind="scalar")
def coupncd(settlement: Scalar, maturity: Scalar, frequency: Scalar, basis: Scalar = 0) -> float:
    return float(
        (coupons(settlement, maturity, frequency, basis).next - dt.date(1899, 12, 30)).days
    )


@function("COUPPCD", kind="scalar")
def couppcd(settlement: Scalar, maturity: Scalar, frequency: Scalar, basis: Scalar = 0) -> float:
    schedule = coupons(settlement, maturity, frequency, basis)
    return float((schedule.previous - dt.date(1899, 12, 30)).days)


@function("COUPNUM", kind="scalar")
def coupnum(settlement: Scalar, maturity: Scalar, frequency: Scalar, basis: Scalar = 0) -> float:
    return float(coupons(settlement, maturity, frequency, basis).count)


@function("COUPDAYBS", kind="scalar")
def coupdaybs(settlement: Scalar, maturity: Scalar, frequency: Scalar, basis: Scalar = 0) -> float:
    return coupons(settlement, maturity, frequency, basis).days_since_previous


def _iterated_period(maturity: dt.date, settlement: dt.date, step: int) -> tuple[dt.date, dt.date]:
    """The coupon period around the settlement when each date is the previous one less a period,
    a day lost at the end of a short month staying lost."""
    following, date = maturity, maturity
    while date > settlement:
        following, date = date, add_months(date, -step)
    return date, following


@function("COUPDAYS", kind="scalar")
def coupdays(settlement: Scalar, maturity: Scalar, frequency: Scalar, basis: Scalar = 0) -> float:
    schedule = coupons(settlement, maturity, frequency, basis)
    end = as_date(maturity)
    step = 12 // schedule.frequency
    if schedule.basis != 1 or end != end_of_month(end):
        months = {(end.month - 1 + k * step) % 12 + 1 for k in range(schedule.frequency)}
        if schedule.basis == 1 and end.day == 30 and months & {4, 6, 9, 11}:
            raise UncalculableError("COUPDAYS Actual/Actual with the 30th of a short month")
        return schedule.days_in_period
    # At a month-end maturity Excel measures between dates that are neither the coupon dates it
    # reports nor a plain schedule: the plain one (a short month's loss kept) for most
    # settlements, the reported one once the settlement reaches its previous date and both end
    # the period on the same day. This held for every settlement in 11,500 recorded cases.
    previous, following = _iterated_period(end, schedule.settlement, step)
    reported, plain = (schedule.next - schedule.previous).days, (following - previous).days
    if schedule.previous > previous and schedule.next == following:
        return float(reported)
    if schedule.previous < previous and reported != plain:
        raise UncalculableError("COUPDAYS Actual/Actual with a month-end maturity")
    return float(plain)


@function("COUPDAYSNC", kind="scalar")
def coupdaysnc(settlement: Scalar, maturity: Scalar, frequency: Scalar, basis: Scalar = 0) -> float:
    return coupons(settlement, maturity, frequency, basis).days_to_next


def _price(schedule: Coupons, rate: float, yld: float, redemption: float) -> float:
    n, freq = schedule.count, schedule.frequency
    a, e, dsc = schedule.days_since_previous, schedule.pricing_period, schedule.days_remaining
    coupon = 100 * rate / freq
    if n == 1:
        discount = yld / freq * dsc / e + 1
        return (redemption + coupon) / discount - coupon * a / e
    periods = dsc / e
    return (
        redemption / (1 + yld / freq) ** (n - 1 + periods)
        + naive_sum(coupon / (1 + yld / freq) ** (k - 1 + periods) for k in range(1, n + 1))
        - coupon * a / e
    )


@function("PRICE", kind="scalar")
def price(
    settlement: Scalar,
    maturity: Scalar,
    rate: Scalar,
    yld: Scalar,
    redemption: Scalar,
    frequency: Scalar,
    basis: Scalar = 0,
) -> float:
    schedule = coupons(settlement, maturity, frequency, basis)
    r, y, redeem = to_number(rate), to_number(yld), to_number(redemption)
    require(r >= 0 and y >= 0 and redeem > 0)
    return _price(schedule, r, y, redeem)


@function("YIELD", kind="scalar")
def yield_(
    settlement: Scalar,
    maturity: Scalar,
    rate: Scalar,
    pr: Scalar,
    redemption: Scalar,
    frequency: Scalar,
    basis: Scalar = 0,
) -> float:
    schedule = coupons(settlement, maturity, frequency, basis)
    r, target, redeem = to_number(rate), to_number(pr), to_number(redemption)
    require(r >= 0 and target > 0 and redeem > 0)
    if schedule.count == 1:
        if schedule.basis in (2, 3):
            # In the last coupon period Excel counts days as it does for Actual/Actual.
            schedule = replace(schedule, basis=1)
        a, e = schedule.days_since_previous, schedule.pricing_period
        dsr = _between(schedule.settlement, as_date(maturity), schedule.basis)
        freq = schedule.frequency
        invested = target / 100 + a / e * r / freq
        return ((redeem / 100 + r / freq) - invested) / invested * freq * e / dsr
    return newton(lambda y: _price(schedule, r, y, redeem) - target, r if r > 0 else 0.05)


def _macaulay(schedule: Coupons, rate: float, yld: float) -> float:
    n, freq = schedule.count, schedule.frequency
    shift = schedule.days_remaining / schedule.pricing_period
    values, weighted = 0.0, 0.0
    for k in range(1, n + 1):
        flow = 100 * rate / freq + (100 if k == n else 0)
        present = flow / (1 + yld / freq) ** (k - 1 + shift)
        values += present
        weighted += present * (k - 1 + shift) / freq
    return weighted / values


@function("DURATION", kind="scalar")
def duration(
    settlement: Scalar,
    maturity: Scalar,
    coupon: Scalar,
    yld: Scalar,
    frequency: Scalar,
    basis: Scalar = 0,
) -> float:
    schedule = coupons(settlement, maturity, frequency, basis)
    c, y = to_number(coupon), to_number(yld)
    require(c >= 0 and y >= 0)
    return _macaulay(schedule, c, y)


@function("MDURATION", kind="scalar")
def mduration(
    settlement: Scalar,
    maturity: Scalar,
    coupon: Scalar,
    yld: Scalar,
    frequency: Scalar,
    basis: Scalar = 0,
) -> float:
    schedule = coupons(settlement, maturity, frequency, basis)
    c, y = to_number(coupon), to_number(yld)
    require(c >= 0 and y >= 0)
    return _macaulay(schedule, c, y) / (1 + y / schedule.frequency)


def years_between(start: dt.date, end: dt.date, basis: int) -> float:
    """The fraction of a year from start to end, as the discount securities use."""
    require(start < end)
    if basis == 0:
        return bond_days360(start, end) / 360
    return year_fraction(start, end, basis)


def year_span(settlement: Scalar, maturity: Scalar, basis: Scalar) -> float:
    return years_between(as_date(settlement), as_date(maturity), as_basis(basis))


@function("DISC", kind="scalar")
def disc(
    settlement: Scalar, maturity: Scalar, pr: Scalar, redemption: Scalar, basis: Scalar = 0
) -> float:
    p, redeem = to_number(pr), to_number(redemption)
    require(p > 0 and redeem > 0)
    return (redeem - p) / redeem / year_span(settlement, maturity, basis)


@function("PRICEDISC", kind="scalar")
def pricedisc(
    settlement: Scalar, maturity: Scalar, discount: Scalar, redemption: Scalar, basis: Scalar = 0
) -> float:
    d, redeem = to_number(discount), to_number(redemption)
    require(d > 0 and redeem > 0)
    return redeem - d * redeem * year_span(settlement, maturity, basis)


@function("YIELDDISC", kind="scalar")
def yielddisc(
    settlement: Scalar, maturity: Scalar, pr: Scalar, redemption: Scalar, basis: Scalar = 0
) -> float:
    p, redeem = to_number(pr), to_number(redemption)
    require(p > 0 and redeem > 0)
    return (redeem - p) / p / year_span(settlement, maturity, basis)


@function("INTRATE", kind="scalar")
def intrate(
    settlement: Scalar, maturity: Scalar, investment: Scalar, redemption: Scalar, basis: Scalar = 0
) -> float:
    invested, redeem = to_number(investment), to_number(redemption)
    require(invested > 0 and redeem > 0)
    return (redeem - invested) / invested / year_span(settlement, maturity, basis)


@function("RECEIVED", kind="scalar")
def received(
    settlement: Scalar, maturity: Scalar, investment: Scalar, discount: Scalar, basis: Scalar = 0
) -> float:
    invested, d = to_number(investment), to_number(discount)
    require(invested > 0 and d > 0)
    denominator = 1 - d * year_span(settlement, maturity, basis)
    require(denominator > 0)
    return invested / denominator


def _treasury_days(settlement: Scalar, maturity: Scalar) -> int:
    start, end = as_date(settlement), as_date(maturity)
    days = (end - start).days
    require(0 < days <= 365)
    return days


@function("TBILLPRICE", kind="scalar")
def tbillprice(settlement: Scalar, maturity: Scalar, discount: Scalar) -> float:
    d = to_number(discount)
    require(d > 0)
    result = 100 * (1 - d * _treasury_days(settlement, maturity) / 360)
    require(result > 0)
    return result


@function("TBILLYIELD", kind="scalar")
def tbillyield(settlement: Scalar, maturity: Scalar, pr: Scalar) -> float:
    p = to_number(pr)
    require(p > 0)
    return (100 - p) / p * 360 / _treasury_days(settlement, maturity)


@function("TBILLEQ", kind="scalar")
def tbilleq(settlement: Scalar, maturity: Scalar, discount: Scalar) -> float:
    d = to_number(discount)
    days = _treasury_days(settlement, maturity)
    require(d > 0)
    if days <= 182:
        denominator = 360 - d * days
        require(denominator > 0)
        return 365 * d / denominator
    price_ = 100 * (1 - d * days / 360)
    require(price_ > 0)
    a, b, c = days / 730 - 0.25, days / 365, (price_ - 100) / price_
    discriminant = b * b - 4 * a * c
    require(discriminant >= 0)
    return (-b + discriminant**0.5) / (2 * a)
