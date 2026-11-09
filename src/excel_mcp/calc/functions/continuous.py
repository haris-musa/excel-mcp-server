"""Beta, gamma, chi-square, Weibull and lognormal distributions."""

import math
import statistics

from excel_mcp.calc.functions.distributions import probability
from excel_mcp.calc.registry import function
from excel_mcp.calc.specfun import (
    beta_i,
    gamma_p,
    gamma_q,
    log_beta_inverse,
    solve_increasing,
    solve_unbounded,
)
from excel_mcp.calc.values import NUM, FormulaError, Scalar, to_bool, to_int, to_number

MAX_FREEDOM = 1e10


def require(condition: bool) -> None:
    if not condition:
        raise FormulaError(NUM)


def log_beta(a: float, b: float) -> float:
    return -log_beta_inverse(a, b)


def _bounds(low: Scalar, high: Scalar) -> tuple[float, float]:
    a, b = to_number(low), to_number(high)
    require(a < b)
    return a, b


@function("BETA.DIST", kind="scalar")
def beta_dist(
    x: Scalar,
    alpha: Scalar,
    beta: Scalar,
    cumulative: Scalar,
    low: Scalar = 0,
    high: Scalar = 1,
) -> float:
    value, a, b = to_number(x), to_number(alpha), to_number(beta)
    cdf = to_bool(cumulative)
    lower, upper = _bounds(low, high)
    require(a > 0 and b > 0 and lower <= value <= upper)
    scaled = (value - lower) / (upper - lower)
    if cdf:
        return beta_i(a, b, scaled)
    if scaled in (0, 1):
        require((a if scaled == 0 else b) > 1)
        return 0.0
    log_density = (a - 1) * math.log(scaled) + (b - 1) * math.log1p(-scaled) - log_beta(a, b)
    return math.exp(log_density) / (upper - lower)


@function("BETA.INV", "BETAINV", kind="scalar")
def beta_inv(p: Scalar, alpha: Scalar, beta: Scalar, low: Scalar = 0, high: Scalar = 1) -> float:
    target, a, b = to_number(p), to_number(alpha), to_number(beta)
    lower, upper = _bounds(low, high)
    require(0 < target < 1 and a > 0 and b > 0)
    return lower + solve_increasing(lambda x: beta_i(a, b, x), target, 0, 1) * (upper - lower)


@function("BETADIST", kind="scalar")
def betadist(x: Scalar, alpha: Scalar, beta: Scalar, low: Scalar = 0, high: Scalar = 1) -> float:
    return beta_dist(x, alpha, beta, True, low, high)


def _gamma_arguments(x: Scalar, alpha: Scalar, beta: Scalar) -> tuple[float, float, float]:
    value, a, b = to_number(x), to_number(alpha), to_number(beta)
    require(value >= 0 and a > 0 and b > 0)
    return value, a, b


@function("GAMMA.DIST", "GAMMADIST", kind="scalar")
def gamma_dist(x: Scalar, alpha: Scalar, beta: Scalar, cumulative: Scalar) -> float:
    value, a, b = _gamma_arguments(x, alpha, beta)
    if to_bool(cumulative):
        return gamma_p(a, value / b)
    if value == 0:
        require(a > 1)
        return 0.0
    return math.exp((a - 1) * math.log(value) - value / b - math.lgamma(a) - a * math.log(b))


@function("GAMMA.INV", "GAMMAINV", kind="scalar")
def gamma_inv(p: Scalar, alpha: Scalar, beta: Scalar) -> float:
    target, a, b = to_number(p), to_number(alpha), to_number(beta)
    require(0 <= target < 1 and a > 0 and b > 0)
    if target == 0:
        return 0.0
    return b * solve_unbounded(lambda x: gamma_p(a, x), target, a)


def freedom(value: Scalar | int) -> int:
    degrees = value if isinstance(value, int) else to_int(value)
    require(1 <= degrees <= MAX_FREEDOM)
    return degrees


@function("CHISQ.DIST", kind="scalar")
def chisq_dist(x: Scalar, df: Scalar, cumulative: Scalar) -> float:
    value, k = to_number(x), freedom(df)
    require(value >= 0)
    if to_bool(cumulative):
        return gamma_p(k / 2, value / 2)
    if value == 0:
        require(k != 1)
        return 0.5 if k == 2 else 0.0
    return math.exp(
        (k / 2 - 1) * math.log(value) - value / 2 - (k / 2) * math.log(2) - math.lgamma(k / 2)
    )


@function("CHISQ.DIST.RT", "CHIDIST", kind="scalar")
def chisq_dist_rt(x: Scalar, df: Scalar) -> float:
    value, k = to_number(x), freedom(df)
    require(value >= 0)
    return gamma_q(k / 2, value / 2)


@function("CHISQ.INV", kind="scalar")
def chisq_inv(p: Scalar, df: Scalar) -> float:
    target, k = to_number(p), freedom(df)
    require(0 <= target < 1)
    if target == 0:
        return 0.0
    return 2 * solve_unbounded(lambda x: gamma_p(k / 2, x), target, k / 2)


@function("CHISQ.INV.RT", "CHIINV", kind="scalar")
def chisq_inv_rt(p: Scalar, df: Scalar) -> float:
    target, k = to_number(p), freedom(df)
    require(0 < target <= 1)
    return 2 * solve_unbounded(lambda x: -gamma_q(k / 2, x), -target, k / 2)


@function("WEIBULL.DIST", "WEIBULL", kind="scalar")
def weibull_dist(x: Scalar, alpha: Scalar, beta: Scalar, cumulative: Scalar) -> float:
    value, a, b = _gamma_arguments(x, alpha, beta)
    if to_bool(cumulative):
        return -math.expm1(-((value / b) ** a))
    if value == 0:
        return 0.0
    return a / b * (value / b) ** (a - 1) * math.exp(-((value / b) ** a))


def _lognormal(mean: Scalar, sd: Scalar) -> statistics.NormalDist:
    deviation = to_number(sd)
    require(deviation > 0)
    return statistics.NormalDist(to_number(mean), deviation)


@function("LOGNORM.DIST", kind="scalar")
def lognorm_dist(x: Scalar, mean: Scalar, sd: Scalar, cumulative: Scalar) -> float:
    value = to_number(x)
    normal = _lognormal(mean, sd)
    require(value > 0)
    if to_bool(cumulative):
        return normal.cdf(math.log(value))
    return normal.pdf(math.log(value)) / value


@function("LOGNORMDIST", kind="scalar")
def lognormdist(x: Scalar, mean: Scalar, sd: Scalar) -> float:
    return lognorm_dist(x, mean, sd, True)


@function("LOGNORM.INV", "LOGINV", kind="scalar")
def lognorm_inv(p: Scalar, mean: Scalar, sd: Scalar) -> float:
    normal = _lognormal(mean, sd)
    return math.exp(normal.inv_cdf(probability(p)))
