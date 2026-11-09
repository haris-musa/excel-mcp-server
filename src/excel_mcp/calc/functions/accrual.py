"""Accrued interest, securities that pay at maturity, and French depreciation."""

import calendar
import datetime as dt
import math

from excel_mcp.calc.functions.datetime import actual_year_length, add_months, days360
from excel_mcp.calc.functions.securities import (
    as_basis,
    as_date,
    end_of_month,
    require,
    years_between,
)
from excel_mcp.calc.registry import function
from excel_mcp.calc.values import Scalar, UncalculableError, to_bool, to_int, to_number


@function("ACCRINTM", kind="scalar")
def accrintm(
    issue: Scalar, settlement: Scalar, rate: Scalar, par: Scalar, basis: Scalar = 0
) -> float:
    r, face = to_number(rate), to_number(par)
    issued, settled = as_date(issue), as_date(settlement)
    require(r > 0 and face > 0 and issued <= settled)
    return face * r * (years_between(issued, settled, as_basis(basis)) if issued < settled else 0)


def _maturity_terms(
    settlement: Scalar, maturity: Scalar, issue: Scalar, basis: Scalar
) -> tuple[float, float, float]:
    """The years from issue to maturity, settlement to maturity and issue to settlement.

    The first and the last are measured; the second is what is left of the first, which on a
    30/360 basis is not what counting from the settlement date gives at the end of February.
    On an actual basis the year is that of the span from issue to settlement for all three."""
    issued, settled, matures = as_date(issue), as_date(settlement), as_date(maturity)
    base = as_basis(basis)
    require(issued < settled < matures)
    if base == 1:
        year = actual_year_length(issued, settled)
        return (
            (matures - issued).days / year,
            (matures - settled).days / year,
            (settled - issued).days / year,
        )
    whole, accrued = years_between(issued, matures, base), years_between(issued, settled, base)
    if base in (0, 4):
        return whole, whole - accrued, accrued
    return whole, years_between(settled, matures, base), accrued


@function("PRICEMAT", kind="scalar")
def pricemat(
    settlement: Scalar,
    maturity: Scalar,
    issue: Scalar,
    rate: Scalar,
    yld: Scalar,
    basis: Scalar = 0,
) -> float:
    r, y = to_number(rate), to_number(yld)
    require(r >= 0 and y >= 0)
    issued, remaining, accrued = _maturity_terms(settlement, maturity, issue, basis)
    return (100 + issued * r * 100) / (1 + remaining * y) - accrued * r * 100


@function("YIELDMAT", kind="scalar")
def yieldmat(
    settlement: Scalar,
    maturity: Scalar,
    issue: Scalar,
    rate: Scalar,
    pr: Scalar,
    basis: Scalar = 0,
) -> float:
    r, price = to_number(rate), to_number(pr)
    require(r >= 0 and price > 0)
    issued, remaining, accrued = _maturity_terms(settlement, maturity, issue, basis)
    invested = price / 100 + accrued * r
    return ((1 + issued * r) - invested) / invested / remaining


def _days(start: dt.date, end: dt.date, basis: int) -> float:
    if basis in (0, 4):
        if any(d.month == 2 and d.day >= 28 and d == end_of_month(d) for d in (start, end)):
            raise UncalculableError("30/360 days at the end of February")
        return float(days360(start, end, european=basis == 4))
    return float((end - start).days)


def _period_length(start: dt.date, end: dt.date, basis: int, frequency: int, actual: bool) -> float:
    if actual or basis == 1:
        return float((end - start).days)
    return (365 if basis == 3 else 360) / frequency


def grid_date(anchor: dt.date, count: int, frequency: int) -> dt.date:
    """The coupon date ``count`` periods from the anchor. Schedules at the end of a month are
    Excel's month-end case, which is not reproduced."""
    date = add_months(anchor, count * 12 // frequency)
    if date.day != anchor.day or anchor == end_of_month(anchor):
        raise UncalculableError("coupon dates at the end of a month")
    return date


def coupon_periods(
    anchor: dt.date,
    start: dt.date,
    end: dt.date,
    frequency: int,
    basis: int,
    actual_lengths: bool = False,
    later_by_anchor_period: bool = False,
) -> float:
    """The coupon periods from start to end: whole ones count 1, parts the share of the
    period they cover. Periods are laid out from ``anchor``; ``later_by_anchor_period`` measures
    a part after the anchor against the period that ends at it."""
    total = 0.0
    k = 0
    while grid_date(anchor, k, frequency) > start:
        k -= 1
    while grid_date(anchor, k, frequency) < end:
        period_start, period_end = (
            grid_date(anchor, k, frequency),
            grid_date(anchor, k + 1, frequency),
        )
        covered_from, covered_to = max(start, period_start), min(end, period_end)
        if (covered_from, covered_to) == (period_start, period_end):
            total += 1  # a whole period counts as one, whatever its days
        elif covered_to > covered_from:
            if later_by_anchor_period and period_start >= anchor:
                period_start, period_end = grid_date(anchor, -1, frequency), anchor
            total += _days(covered_from, covered_to, basis) / _period_length(
                period_start, period_end, basis, frequency, actual_lengths
            )
        k += 1
    return total


@function("ACCRINT", kind="scalar")
def accrint(
    issue: Scalar,
    first_interest: Scalar,
    settlement: Scalar,
    rate: Scalar,
    par: Scalar,
    frequency: Scalar,
    basis: Scalar = 0,
    calc_method: Scalar = True,
) -> float:
    issued, first, settled = as_date(issue), as_date(first_interest), as_date(settlement)
    r, face, freq, base = to_number(rate), to_number(par), to_int(frequency), as_basis(basis)
    require(r > 0 and face > 0 and freq in (1, 2, 4) and issued < settled)
    from_issue = to_bool(calc_method)
    if not from_issue and (base in (1, 2, 3) or issued == end_of_month(issued)):
        raise UncalculableError("ACCRINT counting from the first interest date")
    periods = coupon_periods(first, issued, settled, freq, base, later_by_anchor_period=True)
    if not from_issue:
        # Excel leaves out the whole coupon periods before the first one but one.
        periods -= max(0, math.floor(coupon_periods(first, issued, first, freq, base)) - 1)
    return face * r / freq * periods


def _first_period_fraction(purchased: dt.date, first: dt.date, base: int) -> float:
    """The part of a year before the first period ends; on an Actual/Actual basis Excel divides
    by the days of the purchase year."""
    if purchased == first:
        return 1.0
    if base == 1:
        return (first - purchased).days / (366 if calendar.isleap(purchased.year) else 365)
    return years_between(purchased, first, base)


@function("AMORLINC", kind="scalar")
def amorlinc(
    cost: Scalar,
    date_purchased: Scalar,
    first_period: Scalar,
    salvage: Scalar,
    period: Scalar,
    rate: Scalar,
    basis: Scalar = 0,
) -> float:
    c, rest, per, r = to_number(cost), to_number(salvage), to_int(period), to_number(rate)
    purchased, first = as_date(date_purchased), as_date(first_period)
    base = as_basis(basis)
    require(c > 0 and 0 <= rest <= c and per >= 0 and r > 0 and base != 2 and purchased <= first)
    yearly = c * r
    initial = min(_first_period_fraction(purchased, first, base) * yearly, c - rest)
    full_periods = int((c - rest - initial) / yearly)
    if per == 0:
        return initial
    if per <= full_periods:
        return yearly
    if per == full_periods + 1:
        return c - rest - yearly * full_periods - initial
    return 0.0


@function("AMORDEGRC", kind="scalar")
def amordegrc(
    cost: Scalar,
    date_purchased: Scalar,
    first_period: Scalar,
    salvage: Scalar,
    period: Scalar,
    rate: Scalar,
    basis: Scalar = 0,
) -> float:
    c, rest, per, r = to_number(cost), to_number(salvage), to_int(period), to_number(rate)
    purchased, first = as_date(date_purchased), as_date(first_period)
    base = as_basis(basis)
    require(
        c > 0 and 0 <= rest <= c and per >= 0 and 0 < r < 0.5 and base != 2 and purchased <= first
    )
    if c == rest:
        return 0.0
    lifetime, elapsed = 1 / r, _first_period_fraction(purchased, first, base)
    # The first of the two periods that share the rest; a purchase on the last day uses one up.
    last_two = math.ceil(lifetime) - 2 - (purchased == first)
    coefficient = 1.0 if lifetime < 3 else 1.5 if lifetime < 4 else 2.0 if lifetime <= 6 else 2.5
    r *= coefficient
    depreciation = min(math.floor(elapsed * r * c + 0.5), c - rest)
    c -= depreciation
    left = c - rest  # what may still be depreciated; the period that passes it is paid in full
    if per == 0:
        return float(depreciation)
    half = 0
    for period in range(1, per + 1):
        if left < 0:
            depreciation = 0
        elif period < last_two:
            depreciation = math.floor(r * c + 0.5)
            c -= r * c
        elif period == last_two:
            half = depreciation = math.floor(c * 0.5 + 0.5)
        elif period == last_two + 1:
            depreciation = half if left > 0 else 0
        else:
            depreciation = 0
        left -= depreciation
    return float(depreciation)
