"""Complex numbers, written as text such as "3+4i"."""

import math
import re
from collections.abc import Callable
from decimal import Decimal

from excel_mcp.calc.functions.helpers import OMITTED, Omitted, flat
from excel_mcp.calc.registry import function
from excel_mcp.calc.values import (
    DIV0,
    NUM,
    VALUE,
    ExcelError,
    FormulaError,
    Scalar,
    UncalculableError,
    Value,
    is_number,
    number_text,
    plain_number,
    to_number,
)

_NUMBER = r"(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?"
_COMPLEX = re.compile(
    rf"(?P<real>[+-]?{_NUMBER})?(?P<sign>[+-]{{1,2}})?(?P<imag>{_NUMBER})?(?P<unit>[ij])?"
)
_MAX_PLAIN = 21
_MAX_ANGLE = 1e8  # beyond this Excel's sine and cosine lose most of their digits


class Complex:
    """A complex number and the letter its text used for the imaginary unit, if it had one."""

    def __init__(self, value: complex, unit: str | None = None) -> None:
        self.value = value
        self.unit = unit

    def __str__(self) -> str:
        real, imag = self.value.real + 0.0, self.value.imag + 0.0
        unit = self.unit or "i"
        if imag == 0:
            return _text(real)
        mark = ("-" if imag < 0 else "") + unit if abs(imag) == 1 else _text(imag) + unit
        if real == 0:
            return mark
        return _text(real) + (mark if mark[0] == "-" else "+" + mark)


def _text(number: float) -> str:
    """A part of a complex number: plain digits up to 21 characters, as these functions write."""
    fixed = f"{abs(Decimal(f'{number:.15g}')):f}"
    if "." in fixed:
        fixed = fixed.rstrip("0").rstrip(".")
    if len(fixed) <= _MAX_PLAIN or number == 0:
        return ("-" if number < 0 else "") + fixed
    return number_text(number)


def _coefficient(text: str) -> float:
    number = plain_number(text)
    if number is None:
        raise FormulaError(NUM)
    return number


def parse_complex(value: Scalar) -> Complex:
    if isinstance(value, ExcelError):
        raise FormulaError(value)
    if isinstance(value, bool):
        raise FormulaError(VALUE)
    if is_number(value):
        return Complex(complex(value, 0))  # pyright: ignore[reportArgumentType]
    text = "".join("" if value is None else str(value).split())
    if isinstance(value, str) and not text:
        raise FormulaError(NUM)
    match = _COMPLEX.fullmatch(text)
    if match is None:
        raise FormulaError(NUM)
    real, sign, imag, unit = (match[k] for k in ("real", "sign", "imag", "unit"))
    if unit is None:
        if sign or imag:
            raise FormulaError(NUM)
        return Complex(complex(_coefficient(real) if real else 0.0, 0))
    if sign is None:
        return Complex(complex(0, _coefficient(real) if real else 1), unit)
    if len(sign) == 2 and imag is None:
        raise FormulaError(NUM)
    magnitude = _coefficient(imag) if imag else 1.0
    negative = sign.count("-") % 2 == 1
    return Complex(
        complex(_coefficient(real) if real else 0.0, -magnitude if negative else magnitude), unit
    )


def _number(value: Scalar) -> float:
    if isinstance(value, bool):
        raise FormulaError(VALUE)
    return to_number(value)


def _unit(numbers: tuple[Complex, ...], error: ExcelError) -> str | None:
    units = {n.unit for n in numbers if n.unit}
    if len(units) > 1:
        raise FormulaError(error)
    return units.pop() if units else None


def _finish(value: complex, *sources: Complex, conflict: ExcelError = VALUE) -> str:
    if not (math.isfinite(value.real) and math.isfinite(value.imag)):
        raise FormulaError(NUM)
    return str(Complex(value, _unit(sources, conflict)))


def _modulus(z: complex) -> float:
    """sqrt(a*a + b*b) as Excel computes it: no scaling, so huge or tiny parts overflow."""
    squares = z.real * z.real + z.imag * z.imag
    if not math.isfinite(squares):
        raise FormulaError(NUM)
    return math.sqrt(squares)


def _argument(z: complex) -> float:
    return math.atan2(z.imag, z.real)


@function("COMPLEX", kind="scalar")
def complex_(real: Scalar, imaginary: Scalar, suffix: Scalar | Omitted = OMITTED) -> str:
    unit = "i" if isinstance(suffix, Omitted) or suffix in (None, "") else str(suffix)
    if unit not in ("i", "j"):
        raise FormulaError(VALUE)
    return str(Complex(complex(_number(real), _number(imaginary)), unit))


@function("IMREAL", kind="scalar")
def imreal(number: Scalar) -> float:
    return parse_complex(number).value.real


@function("IMAGINARY", kind="scalar")
def imaginary(number: Scalar) -> float:
    return parse_complex(number).value.imag


@function("IMABS", kind="scalar")
def imabs(number: Scalar) -> float:
    return abs(parse_complex(number).value)


@function("IMARGUMENT", kind="scalar")
def imargument(number: Scalar) -> float:
    z = parse_complex(number).value
    if z == 0:
        raise FormulaError(DIV0)
    return _argument(z)


@function("IMCONJUGATE", kind="scalar")
def imconjugate(number: Scalar) -> str:
    z = parse_complex(number)
    return _finish(z.value.conjugate(), z)


def _exp(x: float) -> float:
    try:
        return math.exp(x)
    except OverflowError:
        return math.inf


def _sinh(x: float) -> float:
    return (_exp(x) - _exp(-x)) / 2


def _cosh(x: float) -> float:
    return (_exp(x) + _exp(-x)) / 2


def _sin(z: complex) -> complex:
    return complex(math.sin(z.real) * _cosh(z.imag), math.cos(z.real) * _sinh(z.imag))


def _cos(z: complex) -> complex:
    return complex(math.cos(z.real) * _cosh(z.imag), -math.sin(z.real) * _sinh(z.imag))


def _sinh_z(z: complex) -> complex:
    return complex(math.sinh(z.real) * math.cos(z.imag), math.cosh(z.real) * math.sin(z.imag))


def _cosh_z(z: complex) -> complex:
    return complex(math.cosh(z.real) * math.cos(z.imag), math.sinh(z.real) * math.sin(z.imag))


def _exp_z(z: complex) -> complex:
    scale = _exp(z.real)
    return complex(scale * math.cos(z.imag), scale * math.sin(z.imag))


def _polar_sqrt(z: complex) -> complex:
    root, angle = math.sqrt(_modulus(z)), _argument(z)
    return complex(root * math.cos(angle / 2), root * math.sin(angle / 2))


def _log(z: complex) -> complex:
    radius = abs(z.real) if z.imag == 0 else _modulus(z)
    if radius == 0:
        raise FormulaError(NUM)
    return complex(math.log(radius), _argument(z))


def _unary(name: str, operation: Callable[[complex], complex], trig: bool = False) -> None:
    def call(number: Scalar) -> str:
        z = parse_complex(number)
        if trig and max(abs(z.value.real), abs(z.value.imag)) > _MAX_ANGLE:
            raise UncalculableError("angle too large for Excel's sine and cosine")
        try:
            return _finish(operation(z.value), z)
        except (ZeroDivisionError, OverflowError):
            raise FormulaError(NUM) from None

    function(name, kind="scalar")(call)


_unary("IMEXP", _exp_z, trig=True)
_unary("IMSQRT", _polar_sqrt)
_unary("IMLN", _log)
_unary("IMLOG10", lambda z: _log(z) / math.log(10))
_unary("IMLOG2", lambda z: _log(z) / math.log(2))
_unary("IMSIN", _sin, trig=True)
_unary("IMCOS", _cos, trig=True)
_unary("IMTAN", lambda z: _divide(_sin(z), _cos(z)), trig=True)
_unary("IMCOT", lambda z: _divide(_cos(z), _sin(z)), trig=True)
_unary("IMSEC", lambda z: _divide(1, _cos(z)), trig=True)
_unary("IMCSC", lambda z: _divide(1, _sin(z)), trig=True)
_unary("IMSINH", _sinh_z, trig=True)
_unary("IMCOSH", _cosh_z, trig=True)
_unary("IMSECH", lambda z: _divide(1, _cosh_z(z)), trig=True)
_unary("IMCSCH", lambda z: _divide(1, _sinh_z(z)), trig=True)


def _divide(a: complex, b: complex) -> complex:
    """``a / b`` worked out from the parts, so that ``z / z`` is exactly 1 on every platform:
    Python's complex division gives different last bits on different C libraries. Both are
    first scaled by a power of two, which is exact and keeps the squares in range."""
    exponent = -math.frexp(max(abs(b.real), abs(b.imag)))[1]
    ar, ai = math.ldexp(a.real, exponent), math.ldexp(a.imag, exponent)
    br, bi = math.ldexp(b.real, exponent), math.ldexp(b.imag, exponent)
    scale = br * br + bi * bi
    return complex((ar * br + ai * bi) / scale, (ai * br - ar * bi) / scale)


@function("IMSUB", kind="scalar")
def imsub(first: Scalar, second: Scalar) -> str:
    a, b = parse_complex(first), parse_complex(second)
    return _finish(a.value - b.value, a, b, conflict=NUM)


@function("IMDIV", kind="scalar")
def imdiv(first: Scalar, second: Scalar) -> str:
    a, b = parse_complex(first), parse_complex(second)
    if b.value == 0:
        raise FormulaError(NUM)
    return _finish(_divide(a.value, b.value), a, b, conflict=NUM)


@function("IMPOWER", kind="scalar")
def impower(number: Scalar, power: Scalar) -> str:
    z, n = parse_complex(number), _number(power)
    if z.value == 0:
        if n <= 0:
            raise FormulaError(NUM)
        return _finish(0j, z)
    scale, angle = _modulus(z.value) ** n, _argument(z.value)
    return _finish(complex(scale * math.cos(n * angle), scale * math.sin(n * angle)), z)


def _operands(args: tuple[Value, ...]) -> list[Complex | None]:
    """The numbers of the arguments; a blank cell is None."""
    found: list[Complex | None] = []
    for item in flat(args):
        if isinstance(item, ExcelError):
            raise FormulaError(item)
        found.append(None if item is None else parse_complex(item))
    return found


@function("IMSUM")
def imsum(*args: Value) -> str:
    numbers = [n for n in _operands(args) if n is not None]
    return _finish(sum((n.value for n in numbers), 0j), *numbers)


@function("IMPRODUCT")
def improduct(*args: Value) -> str:
    operands = _operands(args)
    product = 1 + 0j
    for n in operands:
        product *= 0j if n is None else n.value
    return _finish(product, *[n for n in operands if n is not None])
