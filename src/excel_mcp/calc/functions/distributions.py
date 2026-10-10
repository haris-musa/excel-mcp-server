"""Probability distributions and standardisation."""

import math
import statistics

from excel_mcp.calc.functions.math import naive_sum
from excel_mcp.calc.registry import function
from excel_mcp.calc.values import (
    NUM,
    FormulaError,
    Scalar,
    to_bool,
    to_int,
    to_number,
)


def _normal(mean: float, sd: float) -> statistics.NormalDist:
    if sd <= 0:
        raise FormulaError(NUM)
    return statistics.NormalDist(mean, sd)


@function("NORM.S.DIST", kind="scalar")
def norm_s_dist(z: Scalar, cumulative: Scalar) -> float:
    dist = statistics.NormalDist()
    return dist.cdf(to_number(z)) if to_bool(cumulative) else dist.pdf(to_number(z))


@function("NORM.DIST", kind="scalar")
def norm_dist(x: Scalar, mean: Scalar, sd: Scalar, cumulative: Scalar) -> float:
    dist = _normal(to_number(mean), to_number(sd))
    return dist.cdf(to_number(x)) if to_bool(cumulative) else dist.pdf(to_number(x))


@function("NORM.S.INV", kind="scalar")
def norm_s_inv(p: Scalar) -> float:
    return _normal(0, 1).inv_cdf(_probability(p))


@function("NORM.INV", kind="scalar")
def norm_inv(p: Scalar, mean: Scalar, sd: Scalar) -> float:
    return _normal(to_number(mean), to_number(sd)).inv_cdf(_probability(p))


def _probability(p: Scalar) -> float:
    value = to_number(p)
    if not 0 < value < 1:
        raise FormulaError(NUM)
    return value


@function("STANDARDIZE", kind="scalar")
def standardize(x: Scalar, mean: Scalar, sd: Scalar) -> float:
    deviation = to_number(sd)
    if deviation <= 0:
        raise FormulaError(NUM)
    return (to_number(x) - to_number(mean)) / deviation


@function("BINOM.DIST", kind="scalar")
def binom_dist(successes: Scalar, trials: Scalar, p: Scalar, cumulative: Scalar) -> float:
    k, n, probability = to_int(successes), to_int(trials), to_number(p)
    if not (0 <= k <= n and 0 <= probability <= 1):
        raise FormulaError(NUM)

    def mass(i: int) -> float:
        return math.comb(n, i) * probability**i * (1 - probability) ** (n - i)

    return naive_sum(mass(i) for i in range(k + 1)) if to_bool(cumulative) else mass(k)


@function("POISSON.DIST", kind="scalar")
def poisson_dist(x: Scalar, mean: Scalar, cumulative: Scalar) -> float:
    k, rate = to_int(x), to_number(mean)
    if k < 0 or rate < 0:
        raise FormulaError(NUM)

    def mass(i: int) -> float:
        return math.exp(-rate) * rate**i / math.factorial(i)

    return naive_sum(mass(i) for i in range(k + 1)) if to_bool(cumulative) else mass(k)


@function("EXPON.DIST", kind="scalar")
def expon_dist(x: Scalar, rate: Scalar, cumulative: Scalar) -> float:
    value, lam = to_number(x), to_number(rate)
    if value < 0 or lam <= 0:
        raise FormulaError(NUM)
    if to_bool(cumulative):
        return 1 - math.exp(-lam * value)
    return lam * math.exp(-lam * value)


@function("CONFIDENCE.NORM", kind="scalar")
def confidence_norm(alpha: Scalar, sd: Scalar, size: Scalar) -> float:
    a, deviation, n = to_number(alpha), to_number(sd), to_int(size)
    if not 0 < a < 1 or deviation <= 0 or n < 1:
        raise FormulaError(NUM)
    return _normal(0, 1).inv_cdf(1 - a / 2) * deviation / math.sqrt(n)
