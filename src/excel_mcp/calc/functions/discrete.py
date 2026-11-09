"""Discrete distributions: hypergeometric, negative binomial, binomial inverse and range."""

import math

from excel_mcp.calc.functions.continuous import require
from excel_mcp.calc.functions.helpers import OMITTED, Omitted, as_grid
from excel_mcp.calc.functions.math import naive_sum
from excel_mcp.calc.registry import function
from excel_mcp.calc.values import (
    NA,
    ExcelError,
    FormulaError,
    Scalar,
    UncalculableError,
    Value,
    check_periods,
    is_number,
    to_bool,
    to_int,
    to_number,
)

MAX_POPULATION = 10_000


def _binomial_mass(k: int, n: int, p: float) -> float:
    return math.comb(n, k) * p**k * (1 - p) ** (n - k)


@function("HYPGEOM.DIST", kind="scalar")
def hypgeom_dist(
    sample_successes: Scalar,
    sample_size: Scalar,
    population_successes: Scalar,
    population_size: Scalar,
    cumulative: Scalar,
) -> float:
    s, n, big_s, big_n = (
        to_int(v) for v in (sample_successes, sample_size, population_successes, population_size)
    )
    cdf = to_bool(cumulative)
    require(s >= 0 and 0 <= n <= big_n and 0 <= big_s <= big_n and big_n > 0)
    if big_n > MAX_POPULATION:
        raise UncalculableError("HYPGEOM.DIST of a large population")
    first, last = max(0, n - big_n + big_s), min(n, big_s)

    def mass(i: int) -> float:
        return math.comb(big_s, i) * math.comb(big_n - big_s, n - i) / math.comb(big_n, n)

    if not cdf:
        return mass(s) if first <= s <= last else 0.0
    return naive_sum(mass(i) for i in range(first, min(s, last) + 1))


@function("HYPGEOMDIST", kind="scalar")
def hypgeomdist(
    sample_successes: Scalar,
    sample_size: Scalar,
    population_successes: Scalar,
    population_size: Scalar,
) -> float:
    return hypgeom_dist(sample_successes, sample_size, population_successes, population_size, False)


@function("NEGBINOM.DIST", kind="scalar")
def negbinom_dist(failures: Scalar, successes: Scalar, p: Scalar, cumulative: Scalar) -> float:
    f, s, probability = to_int(failures), to_int(successes), to_number(p)
    cdf = to_bool(cumulative)
    require(f >= 0 and s >= 1 and 0 < probability < 1)
    check_periods(f + s)

    def mass(i: int) -> float:
        return math.comb(i + s - 1, s - 1) * probability**s * (1 - probability) ** i

    return naive_sum(mass(i) for i in range(f + 1)) if cdf else mass(f)


@function("NEGBINOMDIST", kind="scalar")
def negbinomdist(failures: Scalar, successes: Scalar, p: Scalar) -> float:
    return negbinom_dist(failures, successes, p, False)


@function("BINOM.INV", "CRITBINOM", kind="scalar")
def binom_inv(trials: Scalar, p: Scalar, alpha: Scalar) -> float:
    n, probability, target = to_int(trials), to_number(p), to_number(alpha)
    require(n >= 0 and 0 < probability < 1 and 0 < target < 1)
    check_periods(n)
    total = 0.0
    for k in range(n + 1):
        total += _binomial_mass(k, n, probability)
        if total >= target:
            return float(k)
    return float(n)


@function("BINOM.DIST.RANGE", kind="scalar")
def binom_dist_range(
    trials: Scalar, p: Scalar, low: Scalar, high: Scalar | Omitted = OMITTED
) -> float:
    n, probability, first = to_int(trials), to_number(p), to_int(low)
    last = first if isinstance(high, Omitted) else to_int(high)
    require(n >= 0 and 0 <= probability <= 1 and 0 <= first <= last <= n)
    check_periods(n)
    return naive_sum(_binomial_mass(k, n, probability) for k in range(first, last + 1))


@function("PROB", array=(0, 1))
def prob(
    x_range: Value, probabilities: Value, lower: Scalar, upper: Scalar | Omitted = OMITTED
) -> float:
    xs, ps = as_grid(x_range).flat(), as_grid(probabilities).flat()
    if len(xs) != len(ps):
        raise FormulaError(NA)
    for value in (*xs, *ps):
        if isinstance(value, ExcelError):
            raise FormulaError(value)
    weights = [float(w) if is_number(w) else 0.0 for w in ps]  # pyright: ignore[reportArgumentType]
    low = to_number(lower)
    high = low if isinstance(upper, Omitted) else to_number(upper)
    require(-5e-16 <= sum(weights) - 1 <= 5e-15)  # the slack Excel allows around a total of 1
    return naive_sum(
        w
        for x, w in zip(xs, weights, strict=True)
        if is_number(x) and low <= x <= high  # pyright: ignore[reportOperatorIssue]
    )
