"""Special functions behind the probability distributions: incomplete gamma and beta, inverses."""

import math
from collections.abc import Callable

_EPSILON = 1e-15
_TINY = 1e-300
_MAX_TERMS = 100_000
_STIRLING_FROM = 1e6
_NEAR_ZERO = 1e-6


def gamma_p(a: float, x: float) -> float:
    """The regularized lower incomplete gamma function P(a, x)."""
    if x <= 0:
        return 0.0
    if x < a + 1:
        return _gamma_series(a, x)
    return 1 - _gamma_fraction(a, x)


def gamma_q(a: float, x: float) -> float:
    """The regularized upper incomplete gamma function Q(a, x) = 1 - P(a, x)."""
    if x <= 0:
        return 1.0
    if x < a + 1:
        return 1 - _gamma_series(a, x)
    return _gamma_fraction(a, x)


def _gamma_prefactor(a: float, x: float) -> float:
    return math.exp(a * math.log(x) - x - math.lgamma(a))


def _gamma_series(a: float, x: float) -> float:
    term = total = 1 / a
    n = a
    for _ in range(_MAX_TERMS):
        n += 1
        term *= x / n
        total += term
        if abs(term) < abs(total) * _EPSILON:
            break
    return total * _gamma_prefactor(a, x)


def _gamma_fraction(a: float, x: float) -> float:
    b = x + 1 - a
    c = 1 / _TINY
    d = 1 / b
    h = d
    for i in range(1, _MAX_TERMS):
        an = -i * (i - a)
        b += 2
        d = an * d + b
        d = _TINY if abs(d) < _TINY else d
        c = b + an / c
        c = _TINY if abs(c) < _TINY else c
        d = 1 / d
        delta = d * c
        h *= delta
        if abs(delta - 1) < _EPSILON:
            break
    return h * _gamma_prefactor(a, x)


def beta_i(a: float, b: float, x: float, y: float | None = None) -> float:
    """The regularized incomplete beta function I_x(a, b).

    ``y`` is 1 - x where the caller knows it more precisely than the subtraction does.
    """
    y = 1 - x if y is None else y
    if x <= 0:
        return 0.0
    if y <= 0:
        return 1.0
    front = math.exp(log_beta_inverse(a, b) + a * _log_from(x, y) + b * _log_from(y, x))
    first, second = (a + 1) / (a + b + 2), (b + 1) / (a + b + 2)
    # Where x or y is within rounding of 0 it is the argument that kept its digits.
    if x < _NEAR_ZERO and x <= 4 * first:
        direct = True
    elif y < _NEAR_ZERO and y <= 4 * second:
        direct = False
    else:
        direct = x < first
    if direct:
        return front * _beta_fraction(a, b, x) / a
    return 1 - front * _beta_fraction(b, a, y) / b


def _log_from(value: float, other: float) -> float:
    """log(value) for value + other = 1, exact whichever of them is small."""
    return math.log1p(-other) if other < 0.5 else math.log(value)


def log_beta_inverse(a: float, b: float) -> float:
    """-log B(a, b), exact also when a parameter is large and the gamma terms cancel."""
    big, small = max(a, b), min(a, b)
    return log_gamma_rise(big, small) - math.lgamma(small)


def log_gamma_rise(z: float, step: float) -> float:
    """log Gamma(z + step) - log Gamma(z), by Stirling's series where z is large."""
    if z < _STIRLING_FROM:
        return math.lgamma(z + step) - math.lgamma(z)
    top = z + step
    stirling = (z - 0.5) * math.log1p(step / z) + step * math.log(top) - step
    return stirling + 1 / (12 * top) - 1 / (12 * z) - 1 / (360 * top**3) + 1 / (360 * z**3)


def _beta_fraction(a: float, b: float, x: float) -> float:
    qab, qap, qam = a + b, a + 1, a - 1
    c = 1.0
    d = 1 - qab * x / qap
    d = _TINY if abs(d) < _TINY else d
    d = 1 / d
    h = d
    for m in range(1, _MAX_TERMS):
        m2 = 2 * m
        aa = m * (b - m) * x / ((qam + m2) * (a + m2))
        d = 1 + aa * d
        d = _TINY if abs(d) < _TINY else d
        c = 1 + aa / c
        c = _TINY if abs(c) < _TINY else c
        d = 1 / d
        h *= d * c
        aa = -(a + m) * (qab + m) * x / ((a + m2) * (qap + m2))
        d = 1 + aa * d
        d = _TINY if abs(d) < _TINY else d
        c = 1 + aa / c
        c = _TINY if abs(c) < _TINY else c
        d = 1 / d
        delta = d * c
        h *= delta
        if abs(delta - 1) < _EPSILON:
            break
    return h


def solve_increasing(
    function: Callable[[float], float], target: float, low: float, high: float
) -> float:
    """The x in [low, high] where an increasing function reaches target, by bisection."""
    for _ in range(200):
        middle = (low + high) / 2
        if middle in (low, high):
            break
        if function(middle) < target:
            low = middle
        else:
            high = middle
    return (low + high) / 2


def solve_unbounded(function: Callable[[float], float], target: float, start: float = 1.0) -> float:
    """The x > 0 where an increasing function reaches target, bracketed by doubling."""
    low = high = start
    if function(start) < target:
        while function(high) < target:
            low, high = high, high * 2
            if high > 1e300:
                break
    else:
        while function(low) >= target and low > _TINY:
            high, low = low, low / 2
    return solve_increasing(function, target, low, high)
