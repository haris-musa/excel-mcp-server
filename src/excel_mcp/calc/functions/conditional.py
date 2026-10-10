"""Conditional aggregates: COUNTIF, SUMIFS, AVERAGEIF and friends."""

from typing import TYPE_CHECKING

from excel_mcp.calc.criteria import criterion
from excel_mcp.calc.functions.math import naive_sum
from excel_mcp.calc.parser import Node, Ref
from excel_mcp.calc.registry import function
from excel_mcp.calc.values import (
    DIV0,
    VALUE,
    ExcelError,
    FormulaError,
    RefGrid,
    Scalar,
    UncalculableError,
    Value,
    is_number,
    scalar,
)

if TYPE_CHECKING:
    from excel_mcp.calc.engine import Engine


def _reference(value: Value) -> RefGrid:
    if not isinstance(value, RefGrid):
        raise FormulaError(VALUE)
    return value


def _matches(conditions: tuple[Value, ...]) -> list[tuple[int, int]]:
    """Positions where every (range, criterion) pair of ``conditions`` holds."""
    ranges = [_reference(conditions[i]) for i in range(0, len(conditions), 2)]
    tests = [criterion(scalar(conditions[i])) for i in range(1, len(conditions), 2)]
    first = ranges[0]
    if any(r.height != first.height or r.width != first.width for r in ranges):
        raise FormulaError(VALUE)
    return [
        (row, col)
        for row in range(first.height)
        for col in range(first.width)
        if all(test(r.rows[row][col]) for r, test in zip(ranges, tests, strict=True))  # pyright: ignore[reportArgumentType]
    ]


def _selected(grid: Value, positions: list[tuple[int, int]]) -> list[Scalar]:
    cells = _reference(grid)
    return [cells.rows[row][col] for row, col in positions]  # pyright: ignore[reportReturnType]


def _selected_numbers(values: list[Scalar]) -> list[float]:
    for value in values:
        if isinstance(value, ExcelError):
            raise FormulaError(value)
    return [float(v) for v in values if is_number(v)]  # pyright: ignore[reportArgumentType]


@function("COUNTIF")
def countif(cells: Value, crit: Value) -> float:
    return float(len(_matches((cells, crit))))


@function("COUNTIFS")
def countifs(*conditions: Value) -> float:
    if len(conditions) % 2:
        raise UncalculableError("COUNTIFS: wrong number of arguments")
    return float(len(_matches(conditions)))


@function("SUMIFS")
def sumifs(total: Value, *conditions: Value) -> float:
    if len(conditions) % 2 or not conditions:
        raise UncalculableError("SUMIFS: wrong number of arguments")
    values = _selected_numbers(_selected(_same_shape(total, conditions[0]), _matches(conditions)))
    return naive_sum(values)


@function("AVERAGEIFS")
def averageifs(total: Value, *conditions: Value) -> float:
    if len(conditions) % 2 or not conditions:
        raise UncalculableError("AVERAGEIFS: wrong number of arguments")
    values = _selected_numbers(_selected(_same_shape(total, conditions[0]), _matches(conditions)))
    if not values:
        raise FormulaError(DIV0)
    return naive_sum(values) / len(values)


@function("MINIFS")
def minifs(total: Value, *conditions: Value) -> float:
    if len(conditions) % 2 or not conditions:
        raise UncalculableError("MINIFS: wrong number of arguments")
    values = _selected_numbers(_selected(_same_shape(total, conditions[0]), _matches(conditions)))
    return min(values, default=0.0)


@function("MAXIFS")
def maxifs(total: Value, *conditions: Value) -> float:
    if len(conditions) % 2 or not conditions:
        raise UncalculableError("MAXIFS: wrong number of arguments")
    values = _selected_numbers(_selected(_same_shape(total, conditions[0]), _matches(conditions)))
    return max(values, default=0.0)


def _same_shape(total: Value, other: Value) -> Value:
    first, second = _reference(total), _reference(other)
    if first.height != second.height or first.width != second.width:
        raise FormulaError(VALUE)
    return first


def _resized(engine: "Engine", node: Node | None, cells: RefGrid) -> RefGrid:
    """The range ``node`` names, resized to ``cells``' shape from its top-left cell (SUMIF)."""
    if node is None:
        return cells
    if not isinstance(node, Ref) or node.top is None or node.left is None:
        shown = _reference(engine.eval(node))
        if shown.height == cells.height and shown.width == cells.width:
            return shown
        raise UncalculableError("SUMIF with a computed sum range of a different size")
    resized = Ref(
        node.sheets,
        node.top,
        node.left,
        node.top + cells.height - 1,
        node.left + cells.width - 1,
        node.relative,
    )
    return engine.reference(resized)


def _conditional(engine: "Engine", cells: Node, crit: Node, total: Node | None) -> list[Scalar]:
    source = _reference(engine.eval(cells))
    test = criterion(engine.scalar(crit))
    target = _resized(engine, total, source)
    return [
        target.rows[row][col]  # pyright: ignore[reportReturnType]
        for row in range(source.height)
        for col in range(source.width)
        if test(source.rows[row][col])  # pyright: ignore[reportArgumentType]
    ]


@function("SUMIF", kind="lazy")
def sumif(engine: "Engine", cells: Node, crit: Node, total: Node | None = None) -> float:
    return naive_sum(_selected_numbers(_conditional(engine, cells, crit, total)))


@function("AVERAGEIF", kind="lazy")
def averageif(engine: "Engine", cells: Node, crit: Node, total: Node | None = None) -> float:
    values = _selected_numbers(_conditional(engine, cells, crit, total))
    if not values:
        raise FormulaError(DIV0)
    return naive_sum(values) / len(values)
