"""More math: reciprocal and inverse hyperbolic trigonometry, Gamma, error, Roman numerals."""

import math
import re
from collections.abc import Callable

from excel_mcp.calc.functions.helpers import numbers
from excel_mcp.calc.functions.math import naive_sum
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
    flush,
    to_int,
    to_number,
    to_text,
)

_ANGLE_LIMIT = 2**27
_MAX_WHOLE = 2**53
_EXP_LIMIT = 709.78  # beyond it EXP overflows and the reciprocal of SINH or COSH is 0


def _unary(
    name: str,
    operation: Callable[[float], float],
    domain: Callable[[float], bool] = lambda n: True,
    *,
    zero_is_div0: bool = False,
) -> None:
    def calculate(number: Scalar) -> float:
        n = to_number(number)
        if zero_is_div0 and n == 0:
            raise FormulaError(DIV0)
        if not domain(n):
            raise FormulaError(NUM)
        return operation(n)

    function(name, kind="scalar")(calculate)


def _inside_angle(n: float) -> bool:
    return abs(n) < _ANGLE_LIMIT


_unary("ASINH", math.asinh)
_unary("ACOSH", math.acosh, lambda n: n >= 1)
_unary("ATANH", math.atanh, lambda n: abs(n) < 1)
_unary("ACOT", lambda n: math.pi / 2 - math.atan(n))
_unary("ACOTH", lambda n: math.atanh(1 / n), lambda n: abs(n) > 1)
_unary("COT", lambda n: 1 / math.tan(n), _inside_angle, zero_is_div0=True)
_unary("COTH", lambda n: 1 / math.tanh(n), zero_is_div0=True)
_unary("CSC", lambda n: 1 / math.sin(n), _inside_angle, zero_is_div0=True)
_unary("CSCH", lambda n: 1 / math.sinh(n) if abs(n) < _EXP_LIMIT else 0.0, zero_is_div0=True)
_unary("SEC", lambda n: 1 / math.cos(n), _inside_angle)
_unary("SECH", lambda n: 1 / math.cosh(n) if abs(n) < _EXP_LIMIT else 0.0)
_unary("GAUSS", lambda n: 0.5 * math.erf(n / math.sqrt(2)))
_unary("PHI", lambda n: math.exp(-n * n / 2) / math.sqrt(2 * math.pi))
_unary("FISHER", math.atanh, lambda n: abs(n) < 1)
_unary("FISHERINV", math.tanh)


@function("GAMMA", kind="scalar")
def gamma(number: Scalar) -> float:
    n = to_number(number)
    if n == 0 or (n < 0 and n == math.floor(n)):
        raise FormulaError(NUM)
    return math.gamma(n)


@function("GAMMALN", "GAMMALN.PRECISE", kind="scalar")
def gammaln(number: Scalar) -> float:
    n = to_number(number)
    if n <= 0:
        raise FormulaError(NUM)
    return math.lgamma(n)


@function("ERF", "ERF.PRECISE", kind="scalar")
def erf(lower: Scalar, upper: Scalar = ...) -> float:  # pyright: ignore[reportArgumentType]
    low = to_number(lower)
    if upper is ...:
        return math.erf(low)
    return math.erf(to_number(upper)) - math.erf(low)


@function("ERFC", "ERFC.PRECISE", kind="scalar")
def erfc(number: Scalar) -> float:
    return math.erfc(to_number(number))


@function("DELTA", kind="scalar")
def delta(first: Scalar, second: Scalar = 0) -> float:
    return float(to_number(first) == to_number(second))


@function("GESTEP", kind="scalar")
def gestep(number: Scalar, step: Scalar = 0) -> float:
    return float(to_number(number) >= to_number(step))


def _precise(number: float, significance: float, round_up: bool) -> float:
    if significance == 0:
        return 0.0
    s = abs(significance)
    return (math.ceil(number / s) if round_up else math.floor(number / s)) * s


@function("CEILING.PRECISE", "ISO.CEILING", kind="scalar")
def ceiling_precise(number: Scalar, significance: Scalar = 1) -> float:
    return _precise(to_number(number), to_number(significance), True)


@function("FLOOR.PRECISE", kind="scalar")
def floor_precise(number: Scalar, significance: Scalar = 1) -> float:
    return _precise(to_number(number), to_number(significance), False)


@function("FACTDOUBLE", kind="scalar")
def factdouble(number: Scalar) -> float:
    n = to_int(number)
    if to_number(number) < -1 or n > 300:
        raise FormulaError(NUM)
    return float(math.prod(range(n, 1, -2)))


@function("COMBINA", kind="scalar")
def combina(number: Scalar, chosen: Scalar) -> float:
    n, k = to_int(number), to_int(chosen)
    if to_number(number) < 0 or to_number(chosen) < 0 or (n == 0 and k > 0) or n + k > 2**31:
        raise FormulaError(NUM)
    if n == 0:
        return 1.0
    return float(math.comb(n + k - 1, k))


@function("PERMUTATIONA", kind="scalar")
def permutationa(number: Scalar, chosen: Scalar) -> float:
    n, k = to_int(number), to_int(chosen)
    if n < 0 or k < 0:
        raise FormulaError(NUM)
    return float(n) ** k


@function("MULTINOMIAL")
def multinomial(*args: Value) -> float:
    values = [int(v) for v in numbers(args)]
    if any(v < 0 for v in values) or sum(values) > 170:
        raise FormulaError(NUM)
    return float(math.factorial(sum(values)) // math.prod(math.factorial(v) for v in values))


@function("SERIESSUM", array=(3,))
def seriessum(x: Value, n: Value, m: Value, coefficients: Value) -> float:
    base, power, step = to_number(x), to_number(n), to_number(m)
    items = numbers([coefficients])
    terms = [apply_binary("^", base, power + i * step) for i in range(len(items))]
    for term in terms:
        if isinstance(term, ExcelError):
            raise FormulaError(term)
    return naive_sum(c * t for c, t in zip(items, terms, strict=True))  # pyright: ignore[reportOperatorIssue]


_DIGITS = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"


@function("BASE", kind="scalar")
def base(number: Scalar, radix: Scalar, min_length: Scalar = 0) -> str:
    n, r, width = to_int(number), to_int(radix), to_int(min_length)
    if not 0 <= n < _MAX_WHOLE or not 2 <= r <= 36 or not 0 <= width <= 255:
        raise FormulaError(NUM)
    digits = ""
    while n:
        n, remainder = divmod(n, r)
        digits = _DIGITS[remainder] + digits
    return (digits or "0").rjust(width, "0")


@function("DECIMAL", kind="scalar")
def decimal(text: Scalar, radix: Scalar) -> float:
    r, body = to_int(radix), to_text(text).strip().upper()
    if not 2 <= r <= 36 or len(body) > 255:
        raise FormulaError(NUM)
    if not body:
        return 0.0
    if any(_DIGITS.find(c) < 0 or _DIGITS.find(c) >= r for c in body):
        raise FormulaError(NUM)
    return float(int(body, r))


_ROMAN = [
    (1000, "M"), (900, "CM"), (500, "D"), (400, "CD"), (100, "C"), (90, "XC"),
    (50, "L"), (40, "XL"), (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I"),
]  # fmt: skip
_CONCISE = {
    1: [(950, "LM"), (450, "LD"), (95, "VC"), (45, "VL")],
    2: [(990, "XM"), (490, "XD"), (99, "IC"), (49, "IL")],
    3: [(995, "VM"), (495, "VD")],
    4: [(999, "IM"), (499, "ID")],
}


@function("ROMAN", kind="scalar")
def roman(number: Scalar, form: Scalar = 0) -> str:
    n, style = to_int(number), to_int(form)
    if not 0 <= n <= 3999 or not 0 <= style <= 4:
        raise FormulaError(VALUE)
    table = sorted(
        _ROMAN + [pair for level in range(1, style + 1) for pair in _CONCISE[level]],
        reverse=True,
    )
    out = ""
    for value, numeral in table:
        while n >= value:
            out += numeral
            n -= value
    return out


_ARABIC = re.compile(r"-?[MDCLXVI]*")
_VALUES = {"M": 1000, "D": 500, "C": 100, "L": 50, "X": 10, "V": 5, "I": 1}


@function("ARABIC", kind="scalar")
def arabic(text: Scalar) -> float:
    body = to_text(text).strip().upper()
    sign = -1 if body.startswith("-") else 1
    body = body.removeprefix("-")
    if not _ARABIC.fullmatch(body):
        raise FormulaError(VALUE)
    total = 0
    for index, letter in enumerate(body):
        value = _VALUES[letter]
        later = _VALUES[body[index + 1]] if index + 1 < len(body) else 0
        total += -value if value < later else value
    return float(sign * flush(total))
