"""Functions that build arrays; a cell that is not array-entered shows the first element."""

from typing import TYPE_CHECKING

from excel_mcp.calc.functions.helpers import as_grid
from excel_mcp.calc.functions.lookup import exact_extent
from excel_mcp.calc.parser import Node, Ref
from excel_mcp.calc.registry import function
from excel_mcp.calc.values import (
    CALC,
    VALUE,
    ExcelError,
    FormulaError,
    Grid,
    Scalar,
    UncalculableError,
    Value,
    is_number,
    scalar,
    to_bool,
    to_int,
    to_number,
)

if TYPE_CHECKING:
    from excel_mcp.calc.engine import Engine

_MAX_ELEMENTS = 100_000


@function("ANCHORARRAY", kind="lazy")
def anchor_array(engine: "Engine", reference: Node) -> Value:
    """``A1#``: the whole range the dynamic array formula in A1 spills into."""
    if not isinstance(reference, Ref):
        raise UncalculableError("spill reference to a computed value")
    return engine.anchored(reference)


@function("SEQUENCE", kind="scalar")
def sequence(rows: Scalar, columns: Scalar = 1, start: Scalar = 1, step: Scalar = 1) -> Grid:
    height, width = to_int(rows), to_int(columns)
    if height < 1 or width < 1:
        raise FormulaError(CALC)
    if height * width > _MAX_ELEMENTS:
        raise UncalculableError("array too large")
    first, delta = to_number(start), to_number(step)
    return Grid([[first + (r * width + c) * delta for c in range(width)] for r in range(height)])


def _sort_key(value: Scalar) -> tuple[int, float | str]:
    if is_number(value):
        return (0, float(value))  # pyright: ignore[reportArgumentType]
    if isinstance(value, str):
        return (1, value.casefold())
    raise UncalculableError("sorting blanks, logicals or errors")


def _sorted_lines(lines: list[list[Value]], keys: list[tuple[list[Scalar], bool]]) -> list[int]:
    order = list(range(len(lines)))
    for column, descending in reversed(keys):
        order.sort(key=lambda i, c=column: _sort_key(c[i]), reverse=descending)
    return order


def _take_lines(grid: Grid, order: list[int], by_column: bool) -> Grid:
    if by_column:
        return Grid([[line[i] for i in order] for line in grid.rows])
    return Grid([list(grid.rows[i]) for i in order])


@function("SORT", array=(0,))
def sort(
    array: Value, sort_index: Value = 1, sort_order: Value = 1, by_column: Value = False
) -> Grid:
    grid = exact_extent(array)
    columns = to_bool(scalar(by_column))
    lines = grid.columns() if columns else grid.rows
    index = to_int(scalar(sort_index))
    if not 1 <= index <= (grid.height if columns else grid.width):
        raise FormulaError(VALUE)
    descending = to_int(scalar(sort_order)) == -1
    if to_int(scalar(sort_order)) not in (1, -1):
        raise FormulaError(VALUE)
    key = grid.rows[index - 1] if columns else [line[index - 1] for line in grid.rows]
    order = _sorted_lines(lines, [(key, descending)])  # pyright: ignore[reportArgumentType]
    return _take_lines(grid, order, columns)


@function("SORTBY", array=True)
def sortby(array: Value, *keys: Value) -> Grid:
    grid = exact_extent(array)
    specs: list[tuple[list[Scalar], bool]] = []
    for position in range(0, len(keys), 2):
        key = exact_extent(keys[position]).flat()
        if len(key) != grid.height or exact_extent(keys[position]).width != 1:
            raise UncalculableError("SORTBY by row or with a different size")
        order = to_int(scalar(keys[position + 1])) if position + 1 < len(keys) else 1
        if order not in (1, -1):
            raise FormulaError(VALUE)
        specs.append((key, order == -1))  # pyright: ignore[reportArgumentType]
    return _take_lines(grid, _sorted_lines(grid.rows, specs), False)  # pyright: ignore[reportArgumentType]


def _identity(value: Scalar) -> tuple[object, ...]:
    if isinstance(value, str):
        return ("s", value.casefold())
    return (type(value).__name__ if isinstance(value, bool) else "n", value)


@function("UNIQUE", array=(0,))
def unique(array: Value, by_column: Value = False, exactly_once: Value = False) -> Grid:
    grid = exact_extent(array)
    columns = to_bool(scalar(by_column))
    lines = grid.columns() if columns else [list(line) for line in grid.rows]
    seen: dict[tuple[object, ...], int] = {}
    for line in lines:
        key = tuple(_identity(v) for v in line)  # pyright: ignore[reportArgumentType]
        seen[key] = seen.get(key, 0) + 1
    chosen, done = [], set()
    for line in lines:
        key = tuple(_identity(v) for v in line)  # pyright: ignore[reportArgumentType]
        if key in done or (to_bool(scalar(exactly_once)) and seen[key] > 1):
            continue
        done.add(key)
        chosen.append(line)
    if not chosen:
        raise FormulaError(CALC)
    return Grid([list(c) for c in zip(*chosen, strict=True)]) if columns else Grid(chosen)


@function("FILTER", kind="raw", array=(0, 1))
def filter_(array: Value, include: Value, if_empty: Value = ...) -> Value:  # pyright: ignore[reportArgumentType]
    for argument in (array, include):
        if isinstance(argument, ExcelError):
            raise FormulaError(argument)
    grid, flags = exact_extent(array), exact_extent(include)
    if flags.width == 1 and flags.height == grid.height:
        keep = [i for i, row in enumerate(flags.rows) if _flag(row[0])]
        result = Grid([list(grid.rows[i]) for i in keep]) if keep else None
    elif flags.height == 1 and flags.width == grid.width:
        keep = [i for i, flag in enumerate(flags.rows[0]) if _flag(flag)]
        result = Grid([[line[i] for i in keep] for line in grid.rows]) if keep else None
    else:
        raise FormulaError(VALUE)
    if result is not None:
        return result
    if if_empty is ...:
        raise FormulaError(CALC)
    return if_empty


def _flag(value: Value) -> bool:
    flag = scalar(value)
    if isinstance(flag, ExcelError):
        raise FormulaError(flag)
    if isinstance(flag, str):
        raise FormulaError(VALUE)
    return bool(flag)


def _positions(indexes: tuple[Value, ...], size: int) -> list[int]:
    chosen = []
    for item in indexes:
        n = to_int(scalar(item))
        if n == 0 or abs(n) > size:
            raise FormulaError(VALUE)
        chosen.append(n - 1 if n > 0 else size + n)
    return chosen


@function("CHOOSEROWS", array=(0,))
def chooserows(array: Value, *indexes: Value) -> Grid:
    grid = exact_extent(array)
    return Grid([list(grid.rows[i]) for i in _positions(indexes, grid.height)])


@function("CHOOSECOLS", array=(0,))
def choosecols(array: Value, *indexes: Value) -> Grid:
    grid = exact_extent(array)
    columns = _positions(indexes, grid.width)
    return Grid([[line[i] for i in columns] for line in grid.rows])


def _slice(size: int, count: Value, drop: bool) -> slice:
    n = to_int(scalar(count))
    if abs(n) >= size and not drop:
        return slice(None)
    if drop:
        return slice(n, None) if n >= 0 else slice(None, size + n)
    return slice(None, n) if n >= 0 else slice(size + n, None)


@function("TAKE", array=(0,))
def take(array: Value, rows: Value, columns: Value = None) -> Grid:
    grid = exact_extent(array)
    if to_int(scalar(rows)) == 0 or (columns is not None and to_int(scalar(columns)) == 0):
        raise FormulaError(CALC)
    kept = grid.rows[_slice(grid.height, rows, drop=False)]
    if columns is not None:
        kept = [line[_slice(grid.width, columns, drop=False)] for line in kept]
    return Grid([list(line) for line in kept])


@function("DROP", array=(0,))
def drop(array: Value, rows: Value, columns: Value = None) -> Grid:
    grid = exact_extent(array)
    kept = grid.rows[_slice(grid.height, rows, drop=True)]
    if columns is not None:
        kept = [line[_slice(grid.width, columns, drop=True)] for line in kept]
    if not kept or not kept[0]:
        raise FormulaError(CALC)
    return Grid([list(line) for line in kept])


@function("VSTACK", array=True)
def vstack(*arrays: Value) -> Grid:
    grids = [as_grid(a) for a in arrays]
    width = max(g.width for g in grids)
    pad: Scalar = ExcelError("#N/A")
    return Grid([list(line) + [pad] * (width - g.width) for g in grids for line in g.rows])


@function("HSTACK", array=True)
def hstack(*arrays: Value) -> Grid:
    grids = [as_grid(a) for a in arrays]
    height = max(g.height for g in grids)
    pad: Scalar = ExcelError("#N/A")
    return Grid(
        [
            [v for g in grids for v in (g.rows[r] if r < g.height else [pad] * g.width)]
            for r in range(height)
        ]
    )
