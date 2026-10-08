"""Argument handling shared by the function categories."""

from collections.abc import Iterable

from excel_mcp.calc.values import (
    NA,
    NUM,
    VALUE,
    ExcelError,
    FormulaError,
    Grid,
    Scalar,
    Value,
    is_number,
    to_number,
)


def numbers(args: Iterable[Value]) -> list[float]:
    """Numbers as SUM reads them: text and booleans in ranges are skipped, typed-in ones count."""
    found: list[float] = []
    for arg in args:
        if isinstance(arg, Grid):
            for value in arg.flat():
                if isinstance(value, ExcelError):
                    raise FormulaError(value)
                if is_number(value):
                    found.append(float(value))  # pyright: ignore[reportArgumentType]
        else:
            found.append(to_number(arg))
    return found


def array_numbers(array: Value) -> list[float]:
    """Numbers of an array argument; text, booleans and blanks are skipped."""
    values = array.flat() if isinstance(array, Grid) else [array]
    for value in values:
        if isinstance(value, ExcelError):
            raise FormulaError(value)
    return [float(v) for v in values if is_number(v)]  # pyright: ignore[reportArgumentType]


def flat(args: Iterable[Value]) -> list[Scalar]:
    values: list[Scalar] = []
    for arg in args:
        values.extend(arg.flat() if isinstance(arg, Grid) else [arg])  # pyright: ignore[reportArgumentType]
    return values


def as_grid(value: Value) -> Grid:
    return value if isinstance(value, Grid) else Grid([[value]])


def require_grid(value: Value) -> Grid:
    if not isinstance(value, Grid):
        raise FormulaError(VALUE)
    return value


def same_shape(first: Grid, *others: Grid) -> None:
    if any(g.height != first.height or g.width != first.width for g in others):
        raise FormulaError(VALUE)


def pairs(xs: Value, ys: Value) -> tuple[list[float], list[float]]:
    """Paired numbers of two equally sized arrays, skipping pairs where either is not a number."""
    first, second = as_grid(xs), as_grid(ys)
    if first.height * first.width != second.height * second.width:
        raise FormulaError(NA)
    left, right = [], []
    for a, b in zip(first.flat(), second.flat(), strict=True):
        for value in (a, b):
            if isinstance(value, ExcelError):
                raise FormulaError(value)
        if is_number(a) and is_number(b):
            left.append(float(a))  # pyright: ignore[reportArgumentType]
            right.append(float(b))  # pyright: ignore[reportArgumentType]
    return left, right


def integer(value: Scalar, minimum: int | None = None) -> int:
    number = int(to_number(value))
    if minimum is not None and number < minimum:
        raise FormulaError(NUM)
    return number
