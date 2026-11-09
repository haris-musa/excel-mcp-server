"""Bessel functions of integer order.

J and Y use their power series in decimal arithmetic wide enough for the cancellation,
and Hankel's expansion for large arguments; I is its positive power series; K is the
integral of exp(-x cosh t) cosh(n t), which the trapezoid rule sums to full precision.
"""

import math
from decimal import Decimal, localcontext

PI = Decimal("3.14159265358979323846264338327950288419716939937510582097494459230781640628620899")
EULER = Decimal(
    "0.57721566490153286060651209008240243104215933593992359880576723488486772677766467"
)
SERIES_LIMIT = 60  # beyond this J and Y use the asymptotic expansion
SERIES_LIMIT_HIGH_ORDER = 2000  # where the order is too high for the expansion
_TERMS = 100_000
_UNDERFLOW = -744.4  # log of the smallest float
MAX_GROWTH = 800  # exp(x) overflows beyond about 710


def _digits(x: float) -> int:
    return 40 + int(1.0 * x)


def _harmonic(count: int) -> Decimal:
    return sum((Decimal(1) / k for k in range(1, count + 1)), Decimal(0))


def bessel_j(x: float, n: int) -> float:
    if x == 0:
        return 1.0 if n == 0 else 0.0
    if x < 0:
        return (-1) ** n * bessel_j(-x, n)
    if x > SERIES_LIMIT and x > n * n:
        return _hankel(x, n)[0]
    return float(_series(x, n)[0])


def bessel_y(x: float, n: int) -> float:
    if x > SERIES_LIMIT and x > n * n:
        return _hankel(x, n)[1]
    return float(_series(x, n)[1])


def _series(x: float, n: int) -> tuple[Decimal, Decimal]:
    """J_n(x) and Y_n(x) from their power series, for 0 < x or x == 0."""
    if x > SERIES_LIMIT_HIGH_ORDER:
        raise OverflowError
    with localcontext() as context:
        context.prec = _digits(x)
        half = Decimal(x) / 2
        quarter = half * half
        term = half**n / math.factorial(n)
        j_sum = Decimal(0)
        y_sum = Decimal(0)
        psi_n1 = -EULER + _harmonic(n)  # psi(n + 1)
        psi_k = -EULER  # psi(k + 1)
        k = 0
        while True:
            j_sum += term
            y_sum += term * (psi_k + psi_n1)
            k += 1
            psi_k += Decimal(1) / k
            psi_n1 += Decimal(1) / (n + k)
            term = -term * quarter / (k * (n + k))
            if abs(term) < Decimal(10) ** (-context.prec + 5) * abs(j_sum or 1) and k > x / 2 + n:
                break
            if k > _TERMS:
                raise OverflowError
        finite = sum(
            (
                Decimal(math.factorial(n - k - 1)) / math.factorial(k) * half ** (2 * k - n)
                for k in range(n)
            ),
            Decimal(0),
        )
        y = (2 * j_sum * half.ln() - finite) / PI - y_sum / PI
        return j_sum, y


def _hankel(x: float, n: int) -> tuple[float, float]:
    """J_n(x) and Y_n(x) for large x from Hankel's asymptotic expansion."""
    mu = 4.0 * n * n
    p = q = 0.0
    term = 1.0
    k = 0
    smallest = math.inf
    while True:
        magnitude = abs(term)
        if magnitude > smallest or magnitude < 1e-17:
            break
        smallest = magnitude
        if k % 2 == 0:
            p += term * (-1) ** (k // 2)
        else:
            q += term * (-1) ** (k // 2)
        k += 1
        term *= (mu - (2 * k - 1) ** 2) / (k * 8 * x)
    chi = _reduced(x - (n / 2 + 0.25) * math.pi)
    scale = math.sqrt(2 / (math.pi * x))
    return (
        scale * (p * math.cos(chi) - q * math.sin(chi)),
        scale * (p * math.sin(chi) + q * math.cos(chi)),
    )


def _reduced(angle: float) -> float:
    """The angle modulo 2 pi, with pi to 80 digits so that huge angles keep their digits."""
    with localcontext() as context:
        context.prec = 120
        turns = (Decimal(angle) / (2 * PI)).to_integral_value(rounding="ROUND_FLOOR")
        return float(Decimal(angle) - turns * 2 * PI)


def bessel_i(x: float, n: int) -> float:
    if x == 0:
        return 1.0 if n == 0 else 0.0
    if x < 0:
        return (-1) ** n * bessel_i(-x, n)
    if x > MAX_GROWTH:
        raise OverflowError
    half = x / 2
    if n * math.log(half) < _UNDERFLOW:  # Excel raises (x / 2) to the n-th power on its own
        raise OverflowError
    term = math.exp(n * math.log(half) - math.lgamma(n + 1))
    total = 0.0
    k = 0
    while True:
        total += term
        k += 1
        term *= half * half / (k * (n + k))
        if term <= total * 1e-17 and k > half:
            return total
        if k > _TERMS:
            raise OverflowError


def bessel_k(x: float, n: int) -> float:
    """K_n(x) for x > 0."""
    step = 1 / (16 + 4 * math.sqrt(x))
    total = 0.5 * math.exp(-x)
    previous = math.inf
    t = step
    while t < 800:
        decay = -x * math.cosh(t)
        term = (math.exp(decay + n * t) + math.exp(decay - n * t)) / 2
        total += term
        if term < previous and term <= total * 1e-18:
            return total * step
        previous = term
        t += step
    raise OverflowError
