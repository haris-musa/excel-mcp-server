"""Student's t and F distributions and the hypothesis tests built on them."""

import math
import statistics

from excel_mcp.calc.functions.continuous import freedom, log_beta, require
from excel_mcp.calc.functions.helpers import OMITTED, Omitted, array_numbers, as_grid, pairs
from excel_mcp.calc.functions.math import naive_sum
from excel_mcp.calc.registry import function
from excel_mcp.calc.specfun import beta_i, gamma_q, log_gamma_rise, solve_unbounded
from excel_mcp.calc.values import (
    DIV0,
    NA,
    ExcelError,
    FormulaError,
    Scalar,
    Value,
    is_number,
    to_bool,
    to_int,
    to_number,
)


def t_tail(x: float, df: float) -> float:
    """P(T > x) for x >= 0."""
    return beta_i(df / 2, 0.5, df / (df + x * x), x * x / (df + x * x)) / 2


def _t_right(x: float, df: float) -> float:
    return t_tail(x, df) if x >= 0 else 1 - t_tail(-x, df)


def _t_inverse_tail(tail: float, df: float) -> float:
    if tail >= 0.5:
        return 0.0
    return solve_unbounded(lambda x: -t_tail(x, df), -tail)


@function("T.DIST", kind="scalar")
def t_dist(x: Scalar, df: Scalar, cumulative: Scalar) -> float:
    value, degrees, cdf = to_number(x), to_int(df), to_bool(cumulative)
    if not cdf and 0 <= degrees < 1:
        raise FormulaError(DIV0)  # where the cumulative form is #NUM!
    k = freedom(degrees)
    if cdf:
        return 1 - _t_right(value, k)
    log_density = (
        log_gamma_rise(k / 2, 0.5)
        - 0.5 * math.log(k * math.pi)
        - (k + 1) / 2 * math.log1p(value * value / k)
    )
    return math.exp(log_density)


@function("T.DIST.2T", kind="scalar")
def t_dist_2t(x: Scalar, df: Scalar) -> float:
    value, k = to_number(x), freedom(df)
    require(value >= 0)
    return 2 * t_tail(value, k)


@function("T.DIST.RT", kind="scalar")
def t_dist_rt(x: Scalar, df: Scalar) -> float:
    return _t_right(to_number(x), freedom(df))


@function("TDIST", kind="scalar")
def tdist(x: Scalar, df: Scalar, tails: Scalar) -> float:
    value, k, count = to_number(x), freedom(df), to_int(tails)
    require(value >= 0 and count in (1, 2))
    return count * t_tail(value, k)


def _cauchy_pole(p: float, df: int) -> None:
    if p == 0 and df == 1:
        raise FormulaError(DIV0)  # Excel divides by tan(pi / 2) here


@function("T.INV", kind="scalar")
def t_inv(p: Scalar, df: Scalar) -> float:
    target, k = to_number(p), freedom(df)
    _cauchy_pole(target, k)
    require(0 < target < 1)
    result = _t_inverse_tail(min(target, 1 - target), k)
    return result if target > 0.5 else -result


@function("T.INV.2T", "TINV", kind="scalar")
def t_inv_2t(p: Scalar, df: Scalar) -> float:
    target, k = to_number(p), freedom(df)
    _cauchy_pole(target, k)
    require(0 < target < 2)
    result = _t_inverse_tail(min(target, 2 - target) / 2, k)
    return result if target <= 1 else -result


@function("CONFIDENCE.T", kind="scalar")
def confidence_t(alpha: Scalar, sd: Scalar, size: Scalar) -> float:
    a, deviation, n = to_number(alpha), to_number(sd), to_int(size)
    require(0 < a < 1 and deviation > 0 and n >= 1)
    if n == 1:
        raise FormulaError(DIV0)
    return _t_inverse_tail(a / 2, n - 1) * deviation / math.sqrt(n)


def _f_left(x: float, d1: float, d2: float) -> float:
    return beta_i(d1 / 2, d2 / 2, d1 * x / (d1 * x + d2), d2 / (d1 * x + d2))


def _f_right(x: float, d1: float, d2: float) -> float:
    return beta_i(d2 / 2, d1 / 2, d2 / (d1 * x + d2), d1 * x / (d1 * x + d2))


@function("F.DIST", kind="scalar")
def f_dist(x: Scalar, df1: Scalar, df2: Scalar, cumulative: Scalar) -> float:
    value, d1, d2 = to_number(x), freedom(df1), freedom(df2)
    require(value >= 0)
    if to_bool(cumulative):
        return _f_left(value, d1, d2)
    if value == 0:
        require(d1 != 1)
        return 1.0 if d1 == 2 else 0.0
    log_density = (
        (
            d1 * math.log(d1 * value)
            - d2 * math.log1p(d1 * value / d2)
            - d1 * math.log(d1 * value + d2)
        )
        / 2
        - math.log(value)
        - log_beta(d1 / 2, d2 / 2)
    )
    return math.exp(log_density)


@function("F.DIST.RT", "FDIST", kind="scalar")
def f_dist_rt(x: Scalar, df1: Scalar, df2: Scalar) -> float:
    value, d1, d2 = to_number(x), freedom(df1), freedom(df2)
    require(value >= 0)
    return _f_right(value, d1, d2)


@function("F.INV", kind="scalar")
def f_inv(p: Scalar, df1: Scalar, df2: Scalar) -> float:
    target, d1, d2 = to_number(p), freedom(df1), freedom(df2)
    require(0 <= target < 1)
    if target == 0:
        return 0.0
    return solve_unbounded(lambda x: _f_left(x, d1, d2), target)


@function("F.INV.RT", "FINV", kind="scalar")
def f_inv_rt(p: Scalar, df1: Scalar, df2: Scalar) -> float:
    target, d1, d2 = to_number(p), freedom(df1), freedom(df2)
    require(0 < target <= 1)
    return solve_unbounded(lambda x: -_f_right(x, d1, d2), -target)


def _sample(array: Value) -> list[float]:
    values = array_numbers(array)
    if len(values) < 2:
        raise FormulaError(DIV0)
    return values


def _variance(values: list[float]) -> float:
    mean = naive_sum(values) / len(values)
    return naive_sum((v - mean) ** 2 for v in values) / (len(values) - 1)


@function("F.TEST", "FTEST", array=True)
def f_test(first: Value, second: Value) -> float:
    one, two = _sample(first), _sample(second)
    v1, v2 = _variance(one), _variance(two)
    if v1 == 0 or v2 == 0:
        raise FormulaError(DIV0)
    ratio, d1, d2 = v1 / v2, len(one) - 1, len(two) - 1
    return min(1.0, 2 * min(_f_left(ratio, d1, d2), _f_right(ratio, d1, d2)))


@function("T.TEST", "TTEST", array=True)
def t_test(first: Value, second: Value, tails: Value, kind: Value) -> float:
    count, test = to_int(tails), to_int(kind)
    require(count in (1, 2) and test in (1, 2, 3))
    if test == 1:
        one, two = pairs(first, second)
        if len(one) < 2:
            raise FormulaError(DIV0)
        t, df = _paired(one, two)
    else:
        t, df = _independent(_sample(first), _sample(second), equal=test == 2)
    return count * t_tail(abs(t), df)


def _paired(one: list[float], two: list[float]) -> tuple[float, float]:
    differences = [a - b for a, b in zip(one, two, strict=True)]
    spread = _variance(differences)
    if spread == 0:
        raise FormulaError(DIV0)
    n = len(differences)
    return naive_sum(differences) / n / math.sqrt(spread / n), n - 1


def _independent(one: list[float], two: list[float], equal: bool) -> tuple[float, float]:
    n1, n2 = len(one), len(two)
    v1, v2 = _variance(one), _variance(two)
    difference = naive_sum(one) / n1 - naive_sum(two) / n2
    if equal:
        pooled = ((n1 - 1) * v1 + (n2 - 1) * v2) / (n1 + n2 - 2)
        error = math.sqrt(pooled * (1 / n1 + 1 / n2))
        df = n1 + n2 - 2
    else:
        s1, s2 = v1 / n1, v2 / n2
        error = math.sqrt(s1 + s2)
        df = (s1 + s2) ** 2 / (s1 * s1 / (n1 - 1) + s2 * s2 / (n2 - 1)) if error else 1
    if error == 0:
        raise FormulaError(DIV0)
    return difference / error, df


@function("Z.TEST", "ZTEST", array=(0,))
def z_test(array: Value, x: Scalar, sigma: Scalar | Omitted = OMITTED) -> float:
    values = array_numbers(array)
    n = len(values)
    if not n:
        raise FormulaError(NA)
    if n < 2:
        raise FormulaError(DIV0)
    deviation = statistics.stdev(values) if isinstance(sigma, Omitted) else to_number(sigma)
    require(deviation > 0)
    z = (naive_sum(values) / n - to_number(x)) / (deviation / math.sqrt(n))
    return 1 - statistics.NormalDist().cdf(z)


@function("CHISQ.TEST", "CHITEST", array=True)
def chisq_test(actual: Value, expected: Value) -> float:
    observed, wanted = as_grid(actual), as_grid(expected)
    if (observed.height, observed.width) != (wanted.height, wanted.width):
        raise FormulaError(NA)
    rows, columns = observed.height, observed.width
    df = (rows - 1) * (columns - 1) if rows > 1 and columns > 1 else rows * columns - 1
    if df == 0:
        raise FormulaError(NA)
    statistic = 0.0
    for a, e in zip(observed.flat(), wanted.flat(), strict=True):
        for value in (a, e):
            if isinstance(value, ExcelError):
                raise FormulaError(value)
        if is_number(a) and is_number(e):
            if e == 0:
                raise FormulaError(DIV0)
            statistic += (a - e) ** 2 / e  # pyright: ignore[reportOperatorIssue]
    return gamma_q(df / 2, statistic / 2)
