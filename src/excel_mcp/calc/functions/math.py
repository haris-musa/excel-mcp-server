"""Math and trigonometry functions."""

import math
import operator
from collections.abc import Callable, Iterable
from decimal import ROUND_DOWN, ROUND_HALF_UP, ROUND_UP, Decimal
from functools import reduce

from excel_mcp.calc.functions.helpers import as_grid, numbers
from excel_mcp.calc.operators import apply_binary
from excel_mcp.calc.registry import function
from excel_mcp.calc.values import (
    DIV0,
    NUM,
    VALUE,
    ExcelError,
    FormulaError,
    Scalar,
    Value,
    is_number,
    to_int,
    to_number,
)


def naive_sum(values: Iterable[float]) -> float:
    """Add in order without compensation, as Excel does."""
    return reduce(operator.add, values, 0.0)


def fail(error: ExcelError) -> FormulaError:
    return FormulaError(error)


def round_to(number: float, digits: float, rounding: str) -> float:
    """Round like Excel, which works on the 15 significant digits it displays."""
    places = int(digits)
    if places > 300:
        return number
    exact = Decimal(f"{number:.15g}").scaleb(places)
    return float(exact.quantize(Decimal(1), rounding=rounding).scaleb(-places))


@function("SUM")
def sum_(*args: Value) -> float:
    return naive_sum(numbers(args))


@function("PRODUCT")
def product(*args: Value) -> float:
    values = numbers(args)
    return math.prod(values) if values else 0.0


@function("MIN")
def minimum(*args: Value) -> float:
    return min(numbers(args), default=0.0)


@function("MAX")
def maximum(*args: Value) -> float:
    return max(numbers(args), default=0.0)


@function("SUMSQ")
def sumsq(*args: Value) -> float:
    return naive_sum(v * v for v in numbers(args))


@function("SUMPRODUCT", array=True)
def sumproduct(*arrays: Value) -> float:
    grids = [as_grid(a) for a in arrays]
    first = grids[0]
    if any(g.height != first.height or g.width != first.width for g in grids):
        raise fail(VALUE)
    total = 0.0
    for items in zip(*(g.flat() for g in grids), strict=True):
        for item in items:
            if isinstance(item, ExcelError):
                raise fail(item)
        total += math.prod(float(i) if is_number(i) else 0.0 for i in items)  # pyright: ignore[reportArgumentType]
    return total


@function("ROUND", kind="scalar")
def round_(number: Scalar, digits: Scalar) -> float:
    return round_to(to_number(number), to_number(digits), ROUND_HALF_UP)


@function("ROUNDUP", kind="scalar")
def roundup(number: Scalar, digits: Scalar) -> float:
    return round_to(to_number(number), to_number(digits), ROUND_UP)


@function("ROUNDDOWN", kind="scalar")
def rounddown(number: Scalar, digits: Scalar) -> float:
    return round_to(to_number(number), to_number(digits), ROUND_DOWN)


@function("TRUNC", kind="scalar")
def trunc(number: Scalar, digits: Scalar = 0) -> float:
    return round_to(to_number(number), to_number(digits), ROUND_DOWN)


@function("INT", kind="scalar")
def int_(number: Scalar) -> float:
    return float(math.floor(to_number(number)))


@function("MOD", kind="scalar")
def mod(number: Scalar, divisor: Scalar) -> float:
    n, d = to_number(number), to_number(divisor)
    if d == 0:
        raise fail(DIV0)
    return n - d * math.floor(n / d)


@function("QUOTIENT", kind="scalar")
def quotient(number: Scalar, divisor: Scalar) -> float:
    n, d = to_number(number), to_number(divisor)
    if d == 0:
        raise fail(DIV0)
    return float(math.trunc(n / d))


@function("ABS", kind="scalar")
def abs_(number: Scalar) -> float:
    return abs(to_number(number))


@function("SIGN", kind="scalar")
def sign(number: Scalar) -> float:
    n = to_number(number)
    return float((n > 0) - (n < 0))


@function("SQRT", kind="scalar")
def sqrt(number: Scalar) -> float:
    n = to_number(number)
    if n < 0:
        raise fail(NUM)
    return math.sqrt(n)


@function("POWER", kind="scalar")
def power(number: Scalar, exponent: Scalar) -> Scalar:
    return apply_binary("^", number, exponent)


@function("EXP", kind="scalar")
def exp(number: Scalar) -> float:
    return math.exp(to_number(number))


@function("LN", kind="scalar")
def ln(number: Scalar) -> float:
    n = to_number(number)
    if n <= 0:
        raise fail(NUM)
    return math.log(n)


@function("LOG10", kind="scalar")
def log10(number: Scalar) -> float:
    n = to_number(number)
    if n <= 0:
        raise fail(NUM)
    return math.log10(n)


@function("LOG", kind="scalar")
def log(number: Scalar, base: Scalar = 10) -> float:
    n, b = to_number(number), to_number(base)
    if n <= 0 or b <= 0:
        raise fail(NUM)
    if b == 1:
        raise fail(DIV0)
    return math.log(n) / math.log(b)


@function("PI")
def pi() -> float:
    return math.pi


@function("CEILING", kind="scalar")
def ceiling(number: Scalar, significance: Scalar) -> float:
    n, s = to_number(number), to_number(significance)
    if s == 0:
        return 0.0
    if n > 0 > s:
        raise fail(NUM)
    return math.ceil(n / s) * s


@function("FLOOR", kind="scalar")
def floor(number: Scalar, significance: Scalar) -> float:
    n, s = to_number(number), to_number(significance)
    if s == 0:
        raise fail(DIV0)
    if n > 0 > s:
        raise fail(NUM)
    return math.floor(n / s) * s


@function("MROUND", kind="scalar")
def mround(number: Scalar, multiple: Scalar) -> float:
    n, m = to_number(number), to_number(multiple)
    if m == 0:
        return 0.0
    if n * m < 0:
        raise fail(NUM)
    return math.copysign(math.floor(abs(n / m) + 0.5), m) * abs(m)


@function("EVEN", kind="scalar")
def even(number: Scalar) -> float:
    n = to_number(number)
    return math.copysign(math.ceil(abs(n) / 2) * 2, n)


@function("ODD", kind="scalar")
def odd(number: Scalar) -> float:
    n = to_number(number)
    return math.copysign(math.ceil((abs(n) - 1) / 2) * 2 + 1, n)


@function("FACT", kind="scalar")
def fact(number: Scalar) -> float:
    n = to_int(number)
    if n < 0 or n > 170:
        raise fail(NUM)
    return float(math.factorial(n))


@function("COMBIN", kind="scalar")
def combin(number: Scalar, chosen: Scalar) -> float:
    n, k = to_int(number), to_int(chosen)
    if n < 0 or k < 0 or k > n:
        raise fail(NUM)
    return float(math.comb(n, k))


@function("PERMUT", kind="scalar")
def permut(number: Scalar, chosen: Scalar) -> float:
    n, k = to_int(number), to_int(chosen)
    if n < 0 or k < 0 or k > n:
        raise fail(NUM)
    return float(math.perm(n, k))


def _integers(args: tuple[Value, ...]) -> list[int]:
    values = [int(v) for v in numbers(args)]
    if any(v < 0 for v in values):
        raise fail(NUM)
    return values


@function("GCD")
def gcd(*args: Value) -> float:
    return float(math.gcd(*_integers(args)))


@function("LCM")
def lcm(*args: Value) -> float:
    return float(math.lcm(*_integers(args)))


def _trig(
    name: str, operation: Callable[[float], float], domain: Callable[[float], bool] | None = None
) -> None:
    def calculate(number: Scalar) -> float:
        n = to_number(number)
        if domain and not domain(n):
            raise fail(NUM)
        return operation(n)

    function(name, kind="scalar")(calculate)


_trig("SIN", math.sin)
_trig("COS", math.cos)
_trig("TAN", math.tan)
_trig("ASIN", math.asin, lambda n: -1 <= n <= 1)
_trig("ACOS", math.acos, lambda n: -1 <= n <= 1)
_trig("ATAN", math.atan)
_trig("SINH", math.sinh)
_trig("COSH", math.cosh)
_trig("TANH", math.tanh)
_trig("DEGREES", math.degrees)
_trig("RADIANS", math.radians)


@function("ATAN2", kind="scalar")
def atan2(x: Scalar, y: Scalar) -> float:
    a, b = to_number(x), to_number(y)
    if a == 0 and b == 0:
        raise fail(DIV0)
    return math.atan2(b, a)
