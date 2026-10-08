"""Excel's operators on scalars, and their extension to arrays."""

import math
from collections.abc import Callable

from excel_mcp.calc.values import (
    DIV0,
    NA,
    NUM,
    ExcelError,
    FormulaError,
    Grid,
    RefGrid,
    Scalar,
    UncalculableError,
    Value,
    compare,
    to_number,
    to_text,
)

# Excel snaps the result of a final addition or subtraction to zero when it is only
# rounding noise: =0.1+0.2-0.3 is 0, while =(0.1+0.2-0.3)*1 is 5.55E-17.
_SNAP = 2.0**-50


def apply_binary(op: str, left: Scalar, right: Scalar) -> Scalar:
    if isinstance(left, ExcelError):
        return left
    if isinstance(right, ExcelError):
        return right
    try:
        return _BINARY[op](left, right)
    except FormulaError as error:
        return error.error


def snap(op: str, left: Scalar, right: Scalar, result: Scalar) -> Scalar:
    """Apply Excel's zero snapping to the result of the outermost + or - of a formula."""
    if op not in ("+", "-") or not isinstance(result, float) or result == 0:
        return result
    try:
        scale = max(abs(to_number(left)), abs(to_number(right)))
    except FormulaError:
        return result
    return 0.0 if abs(result) < scale * _SNAP else result


def apply_unary(op: str, operand: Scalar) -> Scalar:
    if isinstance(operand, ExcelError) or op == "+":
        return operand
    try:
        return -to_number(operand)
    except FormulaError as error:
        return error.error


def apply_percent(operand: Scalar) -> Scalar:
    if isinstance(operand, ExcelError):
        return operand
    try:
        return to_number(operand) / 100
    except FormulaError as error:
        return error.error


def elementwise(function: Callable[..., Scalar], *values: Value) -> Value:
    """Apply a scalar function to scalars and grids; grids broadcast like Excel's arrays."""
    grids = [value for value in values if isinstance(value, Grid)]
    if not grids:
        return function(*values)
    height = max(grid.height for grid in grids)
    width = max(grid.width for grid in grids)
    if any(
        isinstance(g, RefGrid)
        and (g.clipped_rows or g.clipped_cols)
        and (g.height != height or g.width != width)
        for g in grids
    ):
        raise UncalculableError("whole row or column reference combined with a different size")
    rows: list[list[Value]] = []
    for row in range(height):
        out: list[Value] = []
        for col in range(width):
            elements = [_element(value, row, col) for value in values]
            out.append(NA if any(e is _MISSING for e in elements) else function(*elements))
        rows.append(out)
    return Grid(rows)


_MISSING = ExcelError("missing")


def _element(value: Value, row: int, col: int) -> Scalar:
    if not isinstance(value, Grid):
        return value
    row = 0 if value.height == 1 else row
    col = 0 if value.width == 1 else col
    if row >= value.height or col >= value.width:
        return _MISSING
    return value.rows[row][col]  # pyright: ignore[reportReturnType]


def _arithmetic(function: Callable[[float, float], float]) -> Callable[[Scalar, Scalar], Scalar]:
    def apply(left: Scalar, right: Scalar) -> Scalar:
        result = function(to_number(left), to_number(right))
        return result if math.isfinite(result) else NUM

    return apply


def _divide(left: float, right: float) -> float:
    if right == 0:
        raise FormulaError(DIV0)
    return left / right


def _power(left: float, right: float) -> float:
    if left == 0 and right == 0:
        raise FormulaError(NUM)
    if left == 0 and right < 0:
        raise FormulaError(DIV0)
    if left < 0 and right != int(right):
        return _odd_root(left, right)
    try:
        return left**right
    except OverflowError:
        raise FormulaError(NUM) from None


def _odd_root(base: float, exponent: float) -> float:
    """Excel gives a real result for a negative base raised to 1/n with n odd."""
    n = round(1 / exponent)
    if n % 2 == 0 or abs(1 / exponent - n) > 1e-9:
        raise FormulaError(NUM)
    return -((-base) ** exponent)


def _comparison(test: Callable[[int], bool]) -> Callable[[Scalar, Scalar], Scalar]:
    return lambda left, right: test(compare(left, right))


_BINARY: dict[str, Callable[[Scalar, Scalar], Scalar]] = {
    "+": _arithmetic(lambda a, b: a + b),
    "-": _arithmetic(lambda a, b: a - b),
    "*": _arithmetic(lambda a, b: a * b),
    "/": _arithmetic(_divide),
    "^": _arithmetic(_power),
    "&": lambda left, right: to_text(left) + to_text(right),
    "=": _comparison(lambda c: c == 0),
    "<>": _comparison(lambda c: c != 0),
    "<": _comparison(lambda c: c < 0),
    ">": _comparison(lambda c: c > 0),
    "<=": _comparison(lambda c: c <= 0),
    ">=": _comparison(lambda c: c >= 0),
}
