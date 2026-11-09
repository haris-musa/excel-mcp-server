"""Functions that reshape an array: TRANSPOSE, TOCOL, TOROW, WRAPROWS, WRAPCOLS, EXPAND."""

from excel_mcp.calc.functions.helpers import as_grid
from excel_mcp.calc.registry import function
from excel_mcp.calc.values import (
    CALC,
    NA,
    NUM,
    VALUE,
    ExcelError,
    FormulaError,
    Grid,
    Scalar,
    Value,
    check_array_size,
    scalar,
    to_bool,
    to_int,
)


@function("TRANSPOSE", array=(0,))
def transpose(array: Value) -> Grid:
    grid = as_grid(array)
    return Grid([list(column) for column in zip(*grid.rows, strict=True)])


def _flattened(array: Value, ignore: Value, by_column: Value) -> list[Scalar]:
    grid = as_grid(array)
    mode = to_int(scalar(ignore))
    if mode not in (0, 1, 2, 3):
        raise FormulaError(VALUE)
    cells = (
        [c for column in grid.columns() for c in column]
        if to_bool(scalar(by_column))
        else grid.flat()
    )
    kept = [
        item
        for item in cells
        if not (mode in (1, 3) and item is None)
        and not (mode in (2, 3) and isinstance(item, ExcelError))
    ]
    if not kept:
        raise FormulaError(CALC)
    return [0.0 if item is None else item for item in kept]  # pyright: ignore[reportReturnType]


@function("TOCOL", array=(0,))
def tocol(array: Value, ignore: Value = 0, by_column: Value = False) -> Grid:
    return Grid([[item] for item in _flattened(array, ignore, by_column)])


@function("TOROW", array=(0,))
def torow(array: Value, ignore: Value = 0, by_column: Value = False) -> Grid:
    return Grid([_flattened(array, ignore, by_column)])


def _vector(array: Value) -> list[Scalar]:
    grid = as_grid(array)
    if grid.height != 1 and grid.width != 1:
        raise FormulaError(VALUE)
    return [0.0 if item is None else item for item in grid.flat()]  # pyright: ignore[reportReturnType]


def _wrapped(array: Value, count: Value, pad: Value) -> list[list[Scalar]]:
    items = _vector(array)
    size = to_int(scalar(count))
    if size < 1:
        raise FormulaError(NUM)
    size = min(size, len(items))  # a line is never longer than the data
    filler: Scalar = NA if pad is None else scalar(pad)
    lines = [items[i : i + size] for i in range(0, len(items), size)]
    check_array_size(len(lines), size)
    return [line + [filler] * (size - len(line)) for line in lines]


@function("WRAPROWS", array=(0,))
def wraprows(array: Value, count: Value, pad: Value = None) -> Grid:
    return Grid(_wrapped(array, count, pad))


@function("WRAPCOLS", array=(0,))
def wrapcols(array: Value, count: Value, pad: Value = None) -> Grid:
    return Grid([list(row) for row in zip(*_wrapped(array, count, pad), strict=True)])


@function("EXPAND", array=(0,))
def expand(array: Value, rows: Value, columns: Value = None, pad: Value = None) -> Grid:
    grid = as_grid(array)
    height = grid.height if rows is None else to_int(scalar(rows))
    width = grid.width if columns is None else to_int(scalar(columns))
    if height < grid.height or width < grid.width:
        raise FormulaError(VALUE)
    check_array_size(height, width)
    filler: Scalar = NA if pad is None else scalar(pad)
    return Grid(
        [
            [
                grid.rows[r][c] if r < grid.height and c < grid.width else filler
                for c in range(width)
            ]
            for r in range(height)
        ]
    )
