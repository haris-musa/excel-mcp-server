"""Securities with an odd first or last coupon period."""

import datetime as dt
from dataclasses import dataclass

from excel_mcp.calc.functions.accrual import coupon_periods, grid_date
from excel_mcp.calc.functions.financial import newton
from excel_mcp.calc.functions.securities import as_basis, as_date, require
from excel_mcp.calc.registry import function
from excel_mcp.calc.values import Scalar, UncalculableError, to_int, to_number


@dataclass(frozen=True)
class OddFirst:
    """A security whose first coupon period starts at issue and ends at ``first_coupon``."""

    settlement: dt.date
    maturity: dt.date
    issue: dt.date
    first_coupon: dt.date
    rate: float
    redemption: float
    frequency: int
    basis: int

    def periods(self, start: dt.date, end: dt.date) -> float:
        return coupon_periods(self.first_coupon, start, end, self.frequency, self.basis)

    def price(self, yld: float) -> float:
        coupon = 100 * self.rate / self.frequency
        growth = 1 + yld / self.frequency
        until_first = self.periods(self.settlement, self.first_coupon)
        later = 0
        while grid_date(self.first_coupon, later + 1, self.frequency) <= self.maturity:
            later += 1
        first = coupon * self.periods(self.issue, self.first_coupon)
        total = (first + (0 if later else self.redemption)) / growth**until_first
        for k in range(1, later + 1):
            flow = coupon + (self.redemption if k == later else 0)
            total += flow / growth ** (until_first + k)
        return total - coupon * self.periods(self.issue, self.settlement)


def _odd_first(
    settlement: Scalar,
    maturity: Scalar,
    issue: Scalar,
    first_coupon: Scalar,
    rate: Scalar,
    redemption: Scalar,
    frequency: Scalar,
    basis: Scalar,
) -> OddFirst:
    dates = tuple(as_date(v) for v in (settlement, maturity, issue, first_coupon))
    freq, base = to_int(frequency), as_basis(basis)
    r, redeem = to_number(rate), to_number(redemption)
    require(
        dates[1] > dates[3] > dates[0] > dates[2] and r >= 0 and redeem > 0 and freq in (1, 2, 4)
    )
    after_first = 1
    while grid_date(dates[3], after_first, freq) < dates[1]:
        after_first += 1
    require(grid_date(dates[3], after_first, freq) == dates[1])  # maturity ends a coupon period
    if base in (2, 3) and dates[2] < grid_date(dates[3], -1, freq):
        raise UncalculableError("ODDF with a long first period in a 365-day basis")
    return OddFirst(dates[0], dates[1], dates[2], dates[3], r, redeem, freq, base)


@function("ODDFPRICE", kind="scalar")
def oddfprice(
    settlement: Scalar,
    maturity: Scalar,
    issue: Scalar,
    first_coupon: Scalar,
    rate: Scalar,
    yld: Scalar,
    redemption: Scalar,
    frequency: Scalar,
    basis: Scalar = 0,
) -> float:
    security = _odd_first(
        settlement, maturity, issue, first_coupon, rate, redemption, frequency, basis
    )
    y = to_number(yld)
    require(y >= 0)
    return security.price(y)


@function("ODDFYIELD", kind="scalar")
def oddfyield(
    settlement: Scalar,
    maturity: Scalar,
    issue: Scalar,
    first_coupon: Scalar,
    rate: Scalar,
    pr: Scalar,
    redemption: Scalar,
    frequency: Scalar,
    basis: Scalar = 0,
) -> float:
    security = _odd_first(
        settlement, maturity, issue, first_coupon, rate, redemption, frequency, basis
    )
    target = to_number(pr)
    require(target > 0)
    return newton(
        lambda y: security.price(y) - target, security.rate if security.rate > 0 else 0.05
    )


@dataclass(frozen=True)
class OddLast:
    """A security whose last coupon period runs from ``last_interest`` to maturity."""

    settlement: dt.date
    maturity: dt.date
    last_interest: dt.date
    rate: float
    redemption: float
    frequency: int
    basis: int

    def periods(self, start: dt.date, end: dt.date) -> float:
        return coupon_periods(
            self.last_interest, start, end, self.frequency, self.basis, self.basis in (2, 3)
        )

    @property
    def coupon(self) -> float:
        return 100 * self.rate / self.frequency

    def price(self, yld: float) -> float:
        owed = self.redemption + self.periods(self.last_interest, self.maturity) * self.coupon
        discounted = owed / (
            1 + self.periods(self.settlement, self.maturity) * yld / self.frequency
        )
        return discounted - self.periods(self.last_interest, self.settlement) * self.coupon

    def yield_for(self, price: float) -> float:
        owed = self.redemption + self.periods(self.last_interest, self.maturity) * self.coupon
        accrued = self.periods(self.last_interest, self.settlement) * self.coupon
        return (
            (owed / (price + accrued) - 1)
            * self.frequency
            / self.periods(self.settlement, self.maturity)
        )


def _odd_last(
    settlement: Scalar,
    maturity: Scalar,
    last_interest: Scalar,
    rate: Scalar,
    redemption: Scalar,
    frequency: Scalar,
    basis: Scalar,
) -> OddLast:
    dates = tuple(as_date(v) for v in (settlement, maturity, last_interest))
    freq, base = to_int(frequency), as_basis(basis)
    r, redeem = to_number(rate), to_number(redemption)
    require(dates[1] > dates[0] > dates[2] and r >= 0 and redeem > 0 and freq in (1, 2, 4))
    return OddLast(dates[0], dates[1], dates[2], r, redeem, freq, base)


@function("ODDLPRICE", kind="scalar")
def oddlprice(
    settlement: Scalar,
    maturity: Scalar,
    last_interest: Scalar,
    rate: Scalar,
    yld: Scalar,
    redemption: Scalar,
    frequency: Scalar,
    basis: Scalar = 0,
) -> float:
    security = _odd_last(settlement, maturity, last_interest, rate, redemption, frequency, basis)
    y = to_number(yld)
    require(y >= 0)
    return security.price(y)


@function("ODDLYIELD", kind="scalar")
def oddlyield(
    settlement: Scalar,
    maturity: Scalar,
    last_interest: Scalar,
    rate: Scalar,
    pr: Scalar,
    redemption: Scalar,
    frequency: Scalar,
    basis: Scalar = 0,
) -> float:
    security = _odd_last(settlement, maturity, last_interest, rate, redemption, frequency, basis)
    target = to_number(pr)
    require(target > 0)
    return security.yield_for(target)
