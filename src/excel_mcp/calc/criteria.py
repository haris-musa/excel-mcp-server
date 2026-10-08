"""COUNTIF-style criteria: ">5", "<>x", "a*", 3."""

import re
from collections.abc import Callable

from excel_mcp.calc.values import ERRORS, ExcelError, Scalar, compare, is_number, parse_number
from excel_mcp.calc.wildcard import Wildcard

Test = Callable[[Scalar], bool]

_OPERATOR = re.compile(r"(<=|>=|<>|<|>|=)?(.*)", re.DOTALL)


def has_wildcard(text: str) -> bool:
    return any(character in text for character in "*?~")


def criterion(raw: Scalar) -> Test:
    """A test for cell values, built from a criterion argument."""
    if isinstance(raw, ExcelError):
        return lambda value: value == raw
    if not isinstance(raw, str):
        return _equals(raw)
    operator, operand = _OPERATOR.fullmatch(raw).groups()  # pyright: ignore[reportOptionalMemberAccess]
    operator = operator or "="
    if operator in ("=", "<>"):
        equal = _equals_text(operand)
        return equal if operator == "=" else (lambda value: not equal(value))
    number = parse_number(operand)
    if number is not None:
        return lambda value: is_number(value) and _ordered(operator, compare(value, number))
    return lambda value: isinstance(value, str) and _ordered(operator, compare(value, operand))


def _ordered(operator: str, order: int) -> bool:
    return {"<": order < 0, ">": order > 0, "<=": order <= 0, ">=": order >= 0}[operator]


def _equals(target: Scalar) -> Test:
    if target is None:
        return lambda value: is_number(value) and value == 0
    if isinstance(target, bool):
        return lambda value: isinstance(value, bool) and value == target
    return lambda value: (
        (is_number(value) and value == target)
        or (isinstance(value, str) and parse_number(value) == target)
    )


def _equals_text(operand: str) -> Test:
    if operand == "":
        return lambda value: value is None
    if operand.upper() in ("TRUE", "FALSE"):
        wanted = operand.upper() == "TRUE"
        return lambda value: isinstance(value, bool) and value == wanted
    if operand.upper() in ERRORS:
        return lambda value: isinstance(value, ExcelError) and value.code == operand.upper()
    number = parse_number(operand)
    if number is not None:
        return lambda value: (
            (is_number(value) and value == number)
            or (isinstance(value, str) and parse_number(value) == number)
        )
    pattern = Wildcard(operand)
    return lambda value: isinstance(value, str) and pattern.fullmatch(value)
