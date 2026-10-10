"""Financial functions: time value of money, cash-flow returns and depreciation."""

import math
from collections.abc import Callable
from itertools import pairwise

from excel_mcp.calc.functions.helpers import array_numbers, numbers
from excel_mcp.calc.functions.math import naive_sum
from excel_mcp.calc.registry import function
from excel_mcp.calc.values import (
    DIV0,
    NUM,
    FormulaError,
    Scalar,
    UncalculableError,
    Value,
    check_periods,
    scalar,
    to_int,
    to_number,
)


def _fv(rate: float, nper: float, pmt: float, pv: float, due: int) -> float:
    if rate == 0:
        return -(pv + pmt * nper)
    growth = (1 + rate) ** nper
    return -(pv * growth + pmt * (1 + rate * due) * (growth - 1) / rate)


def _pmt(rate: float, nper: float, pv: float, fv: float, due: int) -> float:
    if nper == 0:
        raise FormulaError(NUM)
    if rate == 0:
        return -(pv + fv) / nper
    growth = (1 + rate) ** nper
    return -(pv * growth + fv) * rate / ((1 + rate * due) * (growth - 1))


def _type(value: Scalar) -> int:
    return 1 if to_number(value) != 0 else 0


@function("FV", kind="scalar")
def fv(rate: Scalar, nper: Scalar, pmt: Scalar, pv: Scalar = 0, type_: Scalar = 0) -> float:
    return _fv(to_number(rate), to_number(nper), to_number(pmt), to_number(pv), _type(type_))


@function("PV", kind="scalar")
def pv(rate: Scalar, nper: Scalar, pmt: Scalar, fv: Scalar = 0, type_: Scalar = 0) -> float:
    r, n, payment, future, due = (
        to_number(rate),
        to_number(nper),
        to_number(pmt),
        to_number(fv),
        _type(type_),
    )
    if r == 0:
        return -(future + payment * n)
    growth = (1 + r) ** n
    return -(future + payment * (1 + r * due) * (growth - 1) / r) / growth


@function("PMT", kind="scalar")
def pmt(rate: Scalar, nper: Scalar, pv: Scalar, fv: Scalar = 0, type_: Scalar = 0) -> float:
    return _pmt(to_number(rate), to_number(nper), to_number(pv), to_number(fv), _type(type_))


@function("NPER", kind="scalar")
def nper(rate: Scalar, pmt: Scalar, pv: Scalar, fv: Scalar = 0, type_: Scalar = 0) -> float:
    r, payment, present, future, due = (
        to_number(rate),
        to_number(pmt),
        to_number(pv),
        to_number(fv),
        _type(type_),
    )
    if r == 0:
        if payment == 0:
            raise FormulaError(NUM)
        return -(present + future) / payment
    boost = payment * (1 + r * due)
    numerator, denominator = boost - future * r, boost + present * r
    if denominator == 0 or numerator / denominator <= 0 or r <= -1:
        raise FormulaError(NUM)
    return math.log(numerator / denominator) / math.log(1 + r)


@function("RATE", kind="scalar")
def rate_(
    nper: Scalar,
    pmt: Scalar,
    pv: Scalar,
    fv: Scalar = 0,
    type_: Scalar = 0,
    guess: Scalar = 0.1,
) -> float:
    n, payment, present, future, due = (
        to_number(nper),
        to_number(pmt),
        to_number(pv),
        to_number(fv),
        _type(type_),
    )

    def balance(r: float) -> float:
        if r == 0:
            return present + payment * n + future
        growth = (1 + r) ** n
        return present * growth + payment * (1 + r * due) * (growth - 1) / r + future

    return newton(balance, to_number(guess))


def newton(function_: Callable[[float], float], start: float) -> float:
    """A root of ``function_`` by Newton's method; no convergence is #NUM! like in Excel."""
    x = start
    for _ in range(100):
        step = 1e-7 * max(1.0, abs(x))
        y = function_(x)
        slope = (function_(x + step) - function_(x - step)) / (2 * step)
        if slope == 0:
            raise FormulaError(NUM)
        shifted = x - y / slope
        if shifted <= -1:
            shifted = (x - 1) / 2
        if abs(shifted - x) < 1e-14 * max(1.0, abs(x)):
            return shifted
        x = shifted
    raise FormulaError(NUM)


@function("IPMT", kind="scalar")
def ipmt(
    rate: Scalar, per: Scalar, nper: Scalar, pv: Scalar, fv: Scalar = 0, type_: Scalar = 0
) -> float:
    return _interest(*_loan(rate, per, nper, pv, fv, type_))


@function("PPMT", kind="scalar")
def ppmt(
    rate: Scalar, per: Scalar, nper: Scalar, pv: Scalar, fv: Scalar = 0, type_: Scalar = 0
) -> float:
    arguments = _loan(rate, per, nper, pv, fv, type_)
    r, _, n, present, future, due = arguments
    return _pmt(r, n, present, future, due) - _interest(*arguments)


def _loan(
    rate: Scalar, per: Scalar, nper: Scalar, pv: Scalar, fv: Scalar, type_: Scalar
) -> tuple[float, float, float, float, float, int]:
    r, p, n = to_number(rate), to_number(per), to_number(nper)
    if p < 1 or p > n:
        raise FormulaError(NUM)
    return r, p, n, to_number(pv), to_number(fv), _type(type_)


def _interest(r: float, per: float, n: float, present: float, future: float, due: int) -> float:
    payment = _pmt(r, n, present, future, due)
    if due:
        if per == 1:
            return 0.0
        return (_fv(r, per - 2, payment, present, 1) - payment) * r
    return _fv(r, per - 1, payment, present, 0) * r


def _cumulative(
    rate: Scalar,
    nper: Scalar,
    pv: Scalar,
    start: Scalar,
    end: Scalar,
    type_: Scalar,
    portion: Callable[..., float],
) -> float:
    r, n, present = to_number(rate), to_number(nper), to_number(pv)
    first, last, kind = to_int(start), to_int(end), to_number(type_)
    if r <= 0 or n <= 0 or present <= 0 or first < 1 or last < first or last > n:
        raise FormulaError(NUM)
    check_periods(last)
    if kind not in (0, 1):
        raise FormulaError(NUM)
    return naive_sum(portion(r, p, n, present, 0.0, int(kind)) for p in range(first, last + 1))


@function("CUMIPMT", kind="scalar")
def cumipmt(
    rate: Scalar, nper: Scalar, pv: Scalar, start: Scalar, end: Scalar, type_: Scalar
) -> float:
    return _cumulative(rate, nper, pv, start, end, type_, _interest)


@function("CUMPRINC", kind="scalar")
def cumprinc(
    rate: Scalar, nper: Scalar, pv: Scalar, start: Scalar, end: Scalar, type_: Scalar
) -> float:
    def principal(r: float, per: float, n: float, present: float, future: float, due: int) -> float:
        return _pmt(r, n, present, future, due) - _interest(r, per, n, present, future, due)

    return _cumulative(rate, nper, pv, start, end, type_, principal)


@function("NPV")
def npv(rate: Value, *values: Value) -> float:
    r = to_number(scalar(rate))
    if r == -1:
        raise FormulaError(DIV0)
    return naive_sum(v / (1 + r) ** i for i, v in enumerate(numbers(values), start=1))


def _single_sign_change(flows: list[float]) -> bool:
    signs = [f > 0 for f in flows if f != 0]
    return sum(a != b for a, b in pairwise(signs)) == 1


def _discount_root(flows: list[float], times: list[float]) -> float:
    """The rate at which the flows have zero present value, given a single sign change."""
    if not _single_sign_change(flows):
        raise UncalculableError("IRR with several sign changes has no single answer")

    def value(rate: float) -> float:
        return naive_sum(f * (1 + rate) ** -t for f, t in zip(flows, times, strict=True))

    low, high = -0.999999999999, 1.0
    while value(high) * value(low) > 0:
        high = high * 2 + 1
        if high > 1e12:
            raise FormulaError(NUM)
    for _ in range(300):
        middle = (low + high) / 2
        if value(low) * value(middle) <= 0:
            high = middle
        else:
            low = middle
    return (low + high) / 2


@function("IRR", array=(0,))
def irr(values: Value, guess: Value = 0.1) -> float:
    flows = array_numbers(values)
    scalar(guess)
    if not flows or not (any(f > 0 for f in flows) and any(f < 0 for f in flows)):
        raise FormulaError(NUM)
    return _discount_root(flows, [float(i) for i in range(len(flows))])


def _dated_flows(values: Value, dates: Value) -> tuple[list[float], list[float]]:
    flows = array_numbers(values)
    serials = array_numbers(dates)
    if len(flows) != len(serials):
        raise FormulaError(NUM)
    if any(int(s) < int(serials[0]) for s in serials):
        raise FormulaError(NUM)
    return flows, [(int(s) - int(serials[0])) / 365 for s in serials]


@function("XIRR")
def xirr(values: Value, dates: Value, guess: Value = 0.1) -> float:
    flows, times = _dated_flows(values, dates)
    scalar(guess)
    if not (any(f > 0 for f in flows) and any(f < 0 for f in flows)):
        raise FormulaError(NUM)
    order = sorted(range(len(times)), key=lambda i: times[i])
    return _discount_root([flows[i] for i in order], [times[i] for i in order])


@function("XNPV")
def xnpv(rate: Value, values: Value, dates: Value) -> float:
    r = to_number(scalar(rate))
    flows, times = _dated_flows(values, dates)
    if r <= 0:
        raise FormulaError(NUM)
    return naive_sum(f / (1 + r) ** t for f, t in zip(flows, times, strict=True))


@function("MIRR", array=(0,))
def mirr(values: Value, finance_rate: Value, reinvest_rate: Value) -> float:
    flows = array_numbers(values)
    finance, reinvest = to_number(scalar(finance_rate)), to_number(scalar(reinvest_rate))
    n = len(flows)
    positives = [f if f > 0 else 0.0 for f in flows]
    negatives = [f if f < 0 else 0.0 for f in flows]
    if not any(positives) or not any(negatives):
        raise FormulaError(DIV0)
    future = naive_sum(p * (1 + reinvest) ** (n - 1 - i) for i, p in enumerate(positives))
    present = naive_sum(m / (1 + finance) ** i for i, m in enumerate(negatives))
    return (future / -present) ** (1 / (n - 1)) - 1


@function("FVSCHEDULE")
def fvschedule(principal: Value, schedule: Value) -> float:
    total = to_number(scalar(principal))
    for rate in numbers([schedule]):
        total *= 1 + rate
    return total


@function("EFFECT", kind="scalar")
def effect(nominal: Scalar, periods: Scalar) -> float:
    r, n = to_number(nominal), to_int(periods)
    if r <= 0 or n < 1:
        raise FormulaError(NUM)
    return (1 + r / n) ** n - 1


@function("NOMINAL", kind="scalar")
def nominal(effective: Scalar, periods: Scalar) -> float:
    r, n = to_number(effective), to_int(periods)
    if r <= 0 or n < 1:
        raise FormulaError(NUM)
    return n * ((1 + r) ** (1 / n) - 1)


@function("PDURATION", kind="scalar")
def pduration(rate: Scalar, pv: Scalar, fv: Scalar) -> float:
    r, present, future = to_number(rate), to_number(pv), to_number(fv)
    if r <= 0 or present <= 0 or future <= 0:
        raise FormulaError(NUM)
    return (math.log(future) - math.log(present)) / math.log(1 + r)


@function("RRI", kind="scalar")
def rri(nper: Scalar, pv: Scalar, fv: Scalar) -> float:
    n, present, future = to_number(nper), to_number(pv), to_number(fv)
    if n <= 0 or present == 0:
        raise FormulaError(NUM)
    ratio = future / present
    if ratio < 0:
        raise FormulaError(NUM)
    return ratio ** (1 / n) - 1


@function("ISPMT", kind="scalar")
def ispmt(rate: Scalar, per: Scalar, nper: Scalar, pv: Scalar) -> float:
    r, p, n, present = to_number(rate), to_number(per), to_number(nper), to_number(pv)
    if n == 0:
        raise FormulaError(DIV0)
    return -present * r * (1 - p / n)


def _fraction_digits(fraction: Scalar) -> tuple[int, int]:
    n = to_int(fraction)
    if n < 0:
        raise FormulaError(NUM)
    if n == 0:
        raise FormulaError(DIV0)
    return n, math.ceil(math.log10(n))


@function("DOLLARDE", kind="scalar")
def dollarde(fractional_dollar: Scalar, fraction: Scalar) -> float:
    n, digits = _fraction_digits(fraction)
    value = to_number(fractional_dollar)
    whole = math.trunc(value)
    return whole + (value - whole) * 10**digits / n


@function("DOLLARFR", kind="scalar")
def dollarfr(decimal_dollar: Scalar, fraction: Scalar) -> float:
    n, digits = _fraction_digits(fraction)
    value = to_number(decimal_dollar)
    whole = math.trunc(value)
    return whole + (value - whole) * n / 10**digits
