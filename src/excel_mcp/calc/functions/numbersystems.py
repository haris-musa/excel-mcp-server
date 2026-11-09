"""Conversions between binary, octal, decimal and hexadecimal, and bitwise functions."""

from types import EllipsisType

from excel_mcp.calc.registry import function
from excel_mcp.calc.values import (
    NUM,
    VALUE,
    FormulaError,
    Scalar,
    to_int,
    to_number,
    to_text,
)

# The bits of a number in each system: negative numbers are written in two's complement.
_BITS = {2: 10, 8: 30, 16: 40}
_NAMES = {2: "BIN", 8: "OCT", 10: "DEC", 16: "HEX"}
_DIGITS = "0123456789ABCDEF"
_BIT_LIMIT = 2**48


def _read(text: Scalar, radix: int) -> int:
    if isinstance(text, bool):
        raise FormulaError(VALUE)
    body = to_text(text).upper()
    if len(body) > 10 or any(_DIGITS.find(c) < 0 or _DIGITS.find(c) >= radix for c in body):
        raise FormulaError(NUM)
    value = int(body, radix) if body else 0
    bits = _BITS[radix]
    return value - 2**bits if len(body) == 10 and value >= 2 ** (bits - 1) else value


def _write(value: int, radix: int, places: Scalar | EllipsisType) -> str:
    bits = _BITS.get(radix)
    if radix == 10:
        return str(value)
    assert bits is not None
    if not -(2 ** (bits - 1)) <= value < 2 ** (bits - 1):
        raise FormulaError(NUM)
    if value < 0:
        value += 2**bits
        places = ...
    digits = ""
    while value:
        value, remainder = divmod(value, radix)
        digits = _DIGITS[remainder] + digits
    digits = digits or "0"
    if places is ...:
        return digits
    return digits.rjust(to_int(places), "0") if len(digits) <= to_int(places) else _no(digits)


def _converter(source: int, target: int) -> None:
    def convert(number: Scalar, places: Scalar = ...) -> Scalar:  # pyright: ignore[reportArgumentType]
        if places is not ... and not 1 <= to_int(places) <= 10:
            raise FormulaError(NUM)
        if source == 10:
            if isinstance(number, bool):
                raise FormulaError(VALUE)
            value = to_int(number)
            if not -(2 ** (_BITS[target] - 1)) <= value < 2 ** (_BITS[target] - 1):
                raise FormulaError(NUM)
        else:
            value = _read(number, source)
        if target == 10:
            return float(value)
        return _write(value, target, places)

    function(f"{_NAMES[source]}2{_NAMES[target]}", kind="scalar")(convert)


for _source in (2, 8, 10, 16):
    for _target in (2, 8, 10, 16):
        if _source != _target:
            _converter(_source, _target)


def _no(_: str) -> str:
    raise FormulaError(NUM)


def _bit_operand(value: Scalar) -> int:
    n = to_number(value)
    if n < 0 or n != int(n) or n >= _BIT_LIMIT:
        raise FormulaError(NUM)
    return int(n)


@function("BITAND", kind="scalar")
def bitand(first: Scalar, second: Scalar) -> float:
    return float(_bit_operand(first) & _bit_operand(second))


@function("BITOR", kind="scalar")
def bitor(first: Scalar, second: Scalar) -> float:
    return float(_bit_operand(first) | _bit_operand(second))


@function("BITXOR", kind="scalar")
def bitxor(first: Scalar, second: Scalar) -> float:
    return float(_bit_operand(first) ^ _bit_operand(second))


def _shifted(number: Scalar, shift: Scalar, left: bool) -> float:
    n, amount = _bit_operand(number), to_int(shift)
    if abs(amount) > 53:
        raise FormulaError(NUM)
    result = n << abs(amount) if left == (amount >= 0) else n >> abs(amount)
    result &= 2**53 - 1
    if result >= _BIT_LIMIT:
        raise FormulaError(NUM)
    return float(result)


@function("BITLSHIFT", kind="scalar")
def bitlshift(number: Scalar, shift: Scalar) -> float:
    return _shifted(number, shift, left=True)


@function("BITRSHIFT", kind="scalar")
def bitrshift(number: Scalar, shift: Scalar) -> float:
    return _shifted(number, shift, left=False)
