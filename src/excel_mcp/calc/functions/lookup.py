"""Lookup and reference functions."""

from typing import TYPE_CHECKING

from excel_mcp.calc.functions import search
from excel_mcp.calc.parser import Node, Ref
from excel_mcp.calc.registry import function
from excel_mcp.calc.values import (
    NA,
    REF,
    VALUE,
    ExcelError,
    FormulaError,
    Grid,
    RefGrid,
    Scalar,
    UncalculableError,
    Value,
    scalar,
    to_bool,
    to_int,
    to_number,
)
from excel_mcp.refs import MAX_COLUMN, MAX_ROW

if TYPE_CHECKING:
    from excel_mcp.calc.engine import Engine


def exact_extent(grid: Value, *, rows: bool = True, cols: bool = True) -> Grid:
    """A grid whose size, in the dimensions asked for, is that of the range it came from."""
    if not isinstance(grid, Grid):
        return Grid([[grid]])
    if isinstance(grid, RefGrid) and ((rows and grid.clipped_rows) or (cols and grid.clipped_cols)):
        raise UncalculableError("whole row or column reference where its size matters")
    return grid


def vector(value: Value) -> list[Scalar]:
    grid = value if isinstance(value, Grid) else Grid([[value]])
    if grid.height != 1 and grid.width != 1:
        raise FormulaError(NA)
    return grid.flat()  # pyright: ignore[reportReturnType]


def _at(table: Grid, row: int, col: int) -> Scalar:
    value = table.rows[row][col]
    return 0.0 if value is None else value  # pyright: ignore[reportReturnType]


def _table_lookup(
    lookup: Scalar, table: Grid, index: Scalar, approximate: Scalar, by_row: bool
) -> Scalar:
    position = to_int(index)
    if position < 1:
        raise FormulaError(VALUE)
    lines = table.rows[0] if by_row else [row[0] for row in table.rows]
    if position > (table.height if by_row else table.width):
        raise FormulaError(REF)
    key = search.lookup_key(lookup)
    if to_bool(approximate):
        found = search.last_not_above(key, lines)  # pyright: ignore[reportArgumentType]
    else:
        found = search.exact(key, lines, wildcards=True)  # pyright: ignore[reportArgumentType]
    if found is None:
        raise FormulaError(NA)
    return _at(table, position - 1, found) if by_row else _at(table, found, position - 1)


@function("VLOOKUP")
def vlookup(lookup: Value, table: Value, column: Value, approximate: Value = True) -> Scalar:
    return _table_lookup(
        scalar(lookup),
        exact_extent(table, rows=False),
        scalar(column),
        scalar(approximate),
        by_row=False,
    )


@function("HLOOKUP")
def hlookup(lookup: Value, table: Value, row: Value, approximate: Value = True) -> Scalar:
    return _table_lookup(
        scalar(lookup),
        exact_extent(table, cols=False),
        scalar(row),
        scalar(approximate),
        by_row=True,
    )


@function("MATCH")
def match(lookup: Value, array: Value, kind: Value = 1) -> float:
    order = to_number(scalar(kind))
    key, values = search.lookup_key(scalar(lookup)), vector(array)
    if order > 0:
        found = search.last_not_above(key, values)
    elif order < 0:
        found = search.first_not_below(key, values)
    else:
        found = search.exact(key, values, wildcards=True)
    if found is None:
        raise FormulaError(NA)
    return float(found + 1)


@function("INDEX", array=(0,))
def index(array: Value, row: Value = None, column: Value = None) -> Value:
    grid = array if isinstance(array, Grid) else Grid([[array]])
    r, c = to_int(scalar(row)), to_int(scalar(column))
    if column is None and grid.height > 1 and grid.width > 1:
        raise FormulaError(REF)
    if column is None and grid.height == 1:
        r, c = 0, r
    if r < 0 or c < 0:
        raise FormulaError(VALUE)
    reference = grid if isinstance(grid, RefGrid) else None
    total_rows = reference.total_rows if reference else grid.height
    total_cols = reference.total_cols if reference else grid.width
    if r > total_rows or c > total_cols:
        raise FormulaError(REF)
    blank_row: list[Scalar] = [None] * grid.width
    rows_ = grid.rows if r == 0 else [grid.rows[r - 1] if r <= grid.height else blank_row]
    picked: list[list[Value]] = [  # pyright: ignore[reportAssignmentType]
        line if c == 0 else [line[c - 1] if c <= len(line) else None] for line in rows_
    ]
    if reference:
        top = reference.top + max(r - 1, 0)
        left = reference.left + max(c - 1, 0)
        return RefGrid(
            picked,
            top,
            left,
            reference.sheets,
            clipped_rows=reference.clipped_rows and r == 0,
            clipped_cols=reference.clipped_cols and c == 0,
            sheet=reference.sheet,
        )
    if len(picked) == 1 and len(picked[0]) == 1:
        return picked[0][0]
    return Grid(picked)


def _x_position(
    lookup: Scalar, values: list[Scalar], match_mode: int, search_mode: int
) -> int | None:
    if match_mode not in (-1, 0, 1, 2) or search_mode not in (1, -1, 2, -2):
        raise FormulaError(VALUE)
    if abs(search_mode) == 2:
        if match_mode == 2:
            raise FormulaError(VALUE)
        return search.bisect(lookup, values, nearest=match_mode, descending=search_mode == -2)
    reverse = search_mode == -1
    if match_mode in (0, 2):
        return search.exact(
            lookup, values, wildcards=match_mode == 2, from_end=reverse, blank_matches=True
        )
    return search.nearest(lookup, values, match_mode, from_end=reverse)


@function("XMATCH", array=(1,))
def xmatch(lookup: Value, array: Value, match_mode: Value = 0, search_mode: Value = 1) -> float:
    found = _x_position(
        scalar(lookup),
        vector(array),
        to_int(scalar(match_mode)),
        to_int(scalar(search_mode)),
    )
    if found is None:
        raise FormulaError(NA)
    return float(found + 1)


@function("XLOOKUP", kind="raw", array=(1, 2))
def xlookup(
    lookup: Value,
    lookup_array: Value,
    return_array: Value,
    if_not_found: Value = ...,  # pyright: ignore[reportArgumentType]
    match_mode: Value = 0,
    search_mode: Value = 1,
) -> Value:
    for argument in (lookup, lookup_array, return_array):
        if isinstance(argument, ExcelError):
            raise FormulaError(argument)
    keys = exact_extent(lookup_array)
    results = exact_extent(return_array)
    values = vector(keys)
    by_row = keys.height == 1 and keys.width > 1
    if (results.width if by_row else results.height) != len(values):
        raise FormulaError(VALUE)
    found = _x_position(
        scalar(lookup), values, to_int(scalar(match_mode)), to_int(scalar(search_mode))
    )
    if found is None:
        if if_not_found is ... or if_not_found is None:
            raise FormulaError(NA)
        return if_not_found
    if by_row:
        column = [line[found] for line in results.rows]
        return Grid([[v] for v in column]) if len(column) > 1 else column[0]
    line = list(results.rows[found])
    return Grid([line]) if len(line) > 1 else line[0]


@function("ROWS")
def rows(array: Value) -> float:
    grid = array if isinstance(array, Grid) else Grid([[array]])
    return float(grid.total_rows if isinstance(grid, RefGrid) else grid.height)


@function("COLUMNS")
def columns(array: Value) -> float:
    grid = array if isinstance(array, Grid) else Grid([[array]])
    return float(grid.total_cols if isinstance(grid, RefGrid) else grid.width)


@function("LOOKUP")
def lookup(value: Value, lookup_vector: Value, result_vector: Value = None) -> Scalar:
    keys = exact_extent(lookup_vector)
    if result_vector is None and keys.height > 1 and keys.width > 1:
        raise UncalculableError("LOOKUP in array form")
    found = search.last_not_above(search.lookup_key(scalar(value)), vector(keys))
    if found is None:
        raise FormulaError(NA)
    results = vector(keys if result_vector is None else exact_extent(result_vector))
    if found >= len(results):
        raise FormulaError(NA)
    return results[found]


def offset_target(
    engine: "Engine",
    reference: Node,
    rows: Node,
    columns: Node,
    height: Node | None = None,
    width: Node | None = None,
) -> Ref:
    """The range OFFSET moves to."""
    if not (
        isinstance(reference, Ref)
        and reference.top
        and reference.left
        and reference.bottom
        and reference.right
    ):
        raise UncalculableError("OFFSET from a name or a whole row or column")
    shift_rows, shift_cols = to_int(engine.scalar(rows)), to_int(engine.scalar(columns))
    tall = reference.bottom - reference.top + 1 if height is None else to_int(engine.scalar(height))
    wide = reference.right - reference.left + 1 if width is None else to_int(engine.scalar(width))
    top, left = reference.top + shift_rows, reference.left + shift_cols
    if tall < 1 or wide < 1 or top < 1 or left < 1:
        raise FormulaError(REF)
    if top + tall - 1 > MAX_ROW or left + wide - 1 > MAX_COLUMN:
        raise FormulaError(REF)
    return Ref(reference.sheets, top, left, top + tall - 1, left + wide - 1, False)


@function("OFFSET", kind="lazy")
def offset(
    engine: "Engine",
    reference: Node,
    rows: Node,
    columns: Node,
    height: Node | None = None,
    width: Node | None = None,
) -> Value:
    return engine.reference(offset_target(engine, reference, rows, columns, height, width))
