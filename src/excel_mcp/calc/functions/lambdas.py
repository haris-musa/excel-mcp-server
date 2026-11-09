"""LAMBDA and the functions that call one: BYROW, BYCOL, MAP, REDUCE, SCAN, MAKEARRAY."""

from dataclasses import dataclass
from typing import TYPE_CHECKING

from excel_mcp.calc.parser import Call, Name, Node
from excel_mcp.calc.registry import function
from excel_mcp.calc.values import (
    CALC,
    NA,
    NAME,
    VALUE,
    ExcelError,
    FormulaError,
    Grid,
    Scalar,
    UncalculableError,
    Value,
    check_array_size,
    to_int,
    to_number,
)

if TYPE_CHECKING:
    from excel_mcp.calc.engine import Engine


@dataclass(frozen=True)
class Lambda:
    parameters: tuple[str, ...]
    body: Node


def lambda_of(engine: "Engine", node: Node) -> Lambda:
    """The LAMBDA written in place, or the one a LET name stands for."""
    if isinstance(node, Call) and node.name == "LAMBDA":
        *names, body = node.args
        if not names or not all(isinstance(name, Name) for name in names):
            raise UncalculableError("LAMBDA with parameters that are not names")
        return Lambda(tuple(name.name.casefold() for name in names), body)  # pyright: ignore[reportAttributeAccessIssue]
    if isinstance(node, Name):
        scope = engine.scope_of(node.name)
        found = scope[node.name.casefold()] if scope else None
        if isinstance(found, Lambda):
            return found
        if scope is None:
            raise FormulaError(NAME)  # Excel does not take a function's name in place of a LAMBDA
    raise UncalculableError("a function other than a LAMBDA was given")


def run(engine: "Engine", lam: Lambda, *arguments: Value) -> Value:
    if len(arguments) != len(lam.parameters):
        raise FormulaError(VALUE)
    engine.scopes.append(dict(zip(lam.parameters, arguments, strict=True)))
    try:
        return engine.eval(lam.body, True)
    finally:
        engine.scopes.pop()


def _single(value: Value) -> Scalar:
    """The one value a calculation must give for each row, column or element."""
    if isinstance(value, Grid):
        if value.height != 1 or value.width != 1:
            raise FormulaError(CALC)
        return value.rows[0][0]
    return value


def _grid(engine: "Engine", node: Node) -> Grid:
    value = engine.eval(node, True)
    if isinstance(value, ExcelError):
        raise FormulaError(value)
    return value if isinstance(value, Grid) else Grid([[value]])


@function("LAMBDA", kind="lazy")
def lambda_(engine: "Engine", *args: Node) -> Value:
    raise FormulaError(CALC)  # a LAMBDA nobody calls


@function("BYROW", kind="lazy")
def byrow(engine: "Engine", array: Node, function_: Node) -> Value:
    source, lam = _grid(engine, array), lambda_of(engine, function_)
    return Grid([[_single(run(engine, lam, Grid([row])))] for row in source.rows])


@function("BYCOL", kind="lazy")
def bycol(engine: "Engine", array: Node, function_: Node) -> Value:
    source, lam = _grid(engine, array), lambda_of(engine, function_)
    return Grid(
        [[_single(run(engine, lam, Grid([[v] for v in column]))) for column in source.columns()]]
    )


@function("MAP", kind="lazy")
def map_(engine: "Engine", *args: Node) -> Value:
    *arrays, function_ = args
    sources = [_grid(engine, array) for array in arrays]
    lam = lambda_of(engine, function_)
    height, width = max(s.height for s in sources), max(s.width for s in sources)

    def element(grid: Grid, r: int, c: int) -> Scalar:
        return (
            grid.rows[r][c] if r < grid.height and c < grid.width else NA
        )  # smaller arrays give #N/A

    return Grid(
        [
            [_single(run(engine, lam, *(element(s, r, c) for s in sources))) for c in range(width)]
            for r in range(height)
        ]
    )


def _fold(engine: "Engine", args: tuple[Node, ...]) -> tuple[Value, list[Value], Lambda]:
    *head, function_ = args
    source = _grid(engine, head[-1])
    values = [v for row in source.rows for v in row]
    lam = lambda_of(engine, function_)
    if len(head) == 2:
        return engine.eval(head[0], True), values, lam
    return values[0], values[1:], lam


@function("REDUCE", kind="lazy")
def reduce_(engine: "Engine", *args: Node) -> Value:
    total, values, lam = _fold(engine, args)
    for value in values:
        total = run(engine, lam, total, value)
    return total


@function("SCAN", kind="lazy")
def scan(engine: "Engine", *args: Node) -> Value:
    total, values, lam = _fold(engine, args)
    shape = _grid(engine, args[-2])
    steps = [_single(total)] if len(args) == 2 else []
    for value in values:
        total = _single(run(engine, lam, total, value))
        steps.append(total)
    return Grid([steps[r * shape.width : (r + 1) * shape.width] for r in range(shape.height)])


@function("MAKEARRAY", kind="lazy")
def makearray(engine: "Engine", rows: Node, columns: Node, function_: Node) -> Value:
    counts = engine.scalar(rows), engine.scalar(columns)
    height, width = (to_int(c) for c in counts)
    if height < 1 or width < 1 or any(to_number(c) != to_int(c) for c in counts):
        raise FormulaError(VALUE)
    check_array_size(height, width)
    lam = lambda_of(engine, function_)
    return Grid(
        [
            [_single(run(engine, lam, float(r), float(c))) for c in range(1, width + 1)]
            for r in range(1, height + 1)
        ]
    )


@function("ISOMITTED", kind="lazy")
def isomitted(engine: "Engine", argument: Node) -> Value:
    return False  # every parameter is given: a LAMBDA called with fewer arguments is #VALUE!
