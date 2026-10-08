"""Lookup and reference functions."""

from itertools import pairwise
from typing import TYPE_CHECKING

from excel_mcp.calc.criteria import has_wildcard, wildcard
from excel_mcp.calc.parser import Name, Node, Ref
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
    compare,
    scalar,
    to_bool,
    to_int,
    type_rank,
)

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


def _key(lookup: Scalar) -> Scalar:
    if lookup is None:
        raise UncalculableError("lookup of a blank cell")
    return lookup


def _same_type(value: Scalar, lookup: Scalar) -> bool:
    return not isinstance(value, ExcelError) and type_rank(value) == type_rank(lookup)


def _is_sorted(values: list[Scalar], lookup: Scalar, descending: bool) -> bool:
    typed = [v for v in values if _same_type(v, lookup)]
    return all(compare(a, b) >= 0 if descending else compare(a, b) <= 0 for a, b in pairwise(typed))


def find_position(lookup: Scalar, values: list[Scalar], mode: int, wildcards: bool) -> int | None:
    """0-based index of the match, or None. ``mode``: 0 exact, 1 next smaller, -1 next larger."""
    key = _key(lookup)
    if mode == 0:
        if wildcards and isinstance(key, str) and has_wildcard(key):
            pattern = wildcard(key)
            return next(
                (i for i, v in enumerate(values) if isinstance(v, str) and pattern.fullmatch(v)),
                None,
            )
        return next(
            (i for i, v in enumerate(values) if _same_type(v, key) and compare(v, key) == 0), None
        )
    if not _is_sorted(values, key, descending=mode == -1):
        raise UncalculableError("approximate match on an unsorted range")
    best = None
    for index, value in enumerate(values):
        if not _same_type(value, key):
            continue
        if (compare(value, key) <= 0) if mode == 1 else (compare(value, key) >= 0):
            best = index
        else:
            break
    return best


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
    found = find_position(lookup, lines, 1 if to_bool(approximate) else 0, wildcards=True)  # pyright: ignore[reportArgumentType]
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
    mode = to_int(scalar(kind))
    if mode not in (-1, 0, 1):
        raise UncalculableError("MATCH type")
    found = find_position(scalar(lookup), vector(array), mode, wildcards=True)
    if found is None:
        raise FormulaError(NA)
    return float(found + 1)


@function("INDEX", array=(0,))
def index(array: Value, row: Value = None, column: Value = None) -> Value:
    grid = array if isinstance(array, Grid) else Grid([[array]])
    r, c = to_int(scalar(row)), to_int(scalar(column))
    if column is None and grid.height > 1 and grid.width > 1:
        raise UncalculableError("INDEX of a two-dimensional range with only a row")
    if column is None and grid.height == 1:
        r, c = 0, r
    if r < 0 or c < 0:
        raise FormulaError(VALUE)
    for position, size, clipped in (
        (r, grid.height, isinstance(grid, RefGrid) and grid.clipped_rows),
        (c, grid.width, isinstance(grid, RefGrid) and grid.clipped_cols),
    ):
        if position > size:
            if clipped:
                raise UncalculableError("INDEX beyond the used part of a whole row or column")
            raise FormulaError(REF)
    rows_ = grid.rows if r == 0 else [grid.rows[r - 1]]
    picked = [line if c == 0 else [line[c - 1]] for line in rows_]
    if len(picked) == 1 and len(picked[0]) == 1:
        return picked[0][0]
    if isinstance(grid, RefGrid):
        top = grid.top + max(r - 1, 0)
        left = grid.left + max(c - 1, 0)
        return RefGrid(picked, top, left, grid.sheets)
    return Grid(picked)


def _search_order(size: int, search_mode: int) -> range:
    if search_mode == 1:
        return range(size)
    if search_mode == -1:
        return range(size - 1, -1, -1)
    raise UncalculableError("binary search mode")


def _xmatch(lookup: Scalar, values: list[Scalar], match_mode: int, search_mode: int) -> int | None:
    key = _key(lookup)
    order = _search_order(len(values), search_mode)
    if match_mode in (0, 2):
        if match_mode == 2 and isinstance(key, str):
            pattern = wildcard(key)
            matches = (
                i for i in order if isinstance(text := values[i], str) and pattern.fullmatch(text)
            )
        else:
            matches = (
                i for i in order if _same_type(values[i], key) and compare(values[i], key) == 0
            )
        return next(matches, None)
    if match_mode not in (-1, 1):
        raise UncalculableError("match mode")
    best = None
    for i in order:
        value = values[i]
        if not _same_type(value, key):
            continue
        order_ = compare(value, key)
        if order_ == 0:
            return i
        closer = order_ < 0 if match_mode == -1 else order_ > 0
        if closer and (best is None or compare(value, values[best]) * match_mode < 0):
            best = i
    return best


@function("XMATCH", array=(1,))
def xmatch(lookup: Value, array: Value, match_mode: Value = 0, search_mode: Value = 1) -> float:
    found = _xmatch(
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
    found = _xmatch(scalar(lookup), values, to_int(scalar(match_mode)), to_int(scalar(search_mode)))
    if found is None:
        if if_not_found is ...:
            raise FormulaError(NA)
        return if_not_found
    if by_row:
        column = [line[found] for line in results.rows]
        return Grid([[v] for v in column]) if len(column) > 1 else column[0]
    line = list(results.rows[found])
    return Grid([line]) if len(line) > 1 else line[0]


def _origin(engine: "Engine", node: Node | None) -> Ref:
    if node is None:
        here = engine.here
        return Ref((), here.row, here.col, here.row, here.col, False)
    if isinstance(node, Name):
        raise UncalculableError("ROW or COLUMN of a name")
    if not isinstance(node, Ref) or node.top is None or node.left is None:
        raise UncalculableError("ROW or COLUMN of a computed range")
    return node


@function("ROW", kind="lazy")
def row(engine: "Engine", reference: Node | None = None) -> Value:
    ref = _origin(engine, reference)
    assert ref.top is not None and ref.bottom is not None
    rows = [[float(r)] for r in range(ref.top, ref.bottom + 1)]
    return rows[0][0] if len(rows) == 1 else Grid(rows)  # pyright: ignore[reportReturnType]


@function("COLUMN", kind="lazy")
def column(engine: "Engine", reference: Node | None = None) -> Value:
    ref = _origin(engine, reference)
    assert ref.left is not None and ref.right is not None
    cols = [float(c) for c in range(ref.left, ref.right + 1)]
    return cols[0] if len(cols) == 1 else Grid([cols])


@function("ROWS")
def rows(array: Value) -> float:
    return float(exact_extent(array, cols=False).height)


@function("COLUMNS")
def columns(array: Value) -> float:
    return float(exact_extent(array, rows=False).width)
