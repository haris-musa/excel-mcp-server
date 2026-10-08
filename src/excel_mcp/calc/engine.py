"""Evaluates formulas against the cells of a workbook, the way Excel would.

Anything that cannot be reproduced exactly raises `UncalculableError`; the engine never guesses.
Formulas are evaluated as Excel evaluates ordinary (not array-entered) formulas: a range
where a single value is expected is reduced to the formula cell's own row or column.
"""

import datetime as dt
import logging
import math
from collections.abc import Iterator
from contextlib import suppress
from dataclasses import dataclass
from decimal import Decimal
from typing import Any

from openpyxl import Workbook
from openpyxl.worksheet.formula import ArrayFormula
from openpyxl.worksheet.worksheet import Worksheet

from excel_mcp.calc import functions as functions
from excel_mcp.calc.operators import (
    apply_binary,
    apply_percent,
    apply_unary,
    elementwise,
    snap,
)
from excel_mcp.calc.parser import (
    ArrayLiteral,
    Binary,
    Call,
    ErrorLiteral,
    Literal,
    Name,
    Node,
    Percent,
    Ref,
    Unary,
    parse,
)
from excel_mcp.calc.registry import FUNCTIONS, Function
from excel_mcp.calc.values import (
    ERRORS,
    NUM,
    REF,
    VALUE,
    ExcelError,
    FormulaError,
    Grid,
    RefGrid,
    Scalar,
    UncalculableError,
    Value,
    date_to_serial,
)
from excel_mcp.workspace import worksheets
from excel_mcp.xlfn import FUTURE_FUNCTIONS

logger = logging.getLogger(__name__)

MAX_WORK = 1_000_000
MAX_DEPTH = 30
MAX_NAME_DEPTH = 8

CellKey = tuple[str, int, int]


class TooDeepError(Exception):
    """A chain of dependent cells is too long to follow in one go."""

    def __init__(self, sheet: Worksheet, row: int, col: int) -> None:
        super().__init__("formula chain too deep")
        self.target = (sheet, row, col)


@dataclass
class Position:
    sheet: Worksheet
    row: int
    col: int


class Engine:
    """Calculates formula cells of ``formulas``; ``cached`` holds Excel's stored results."""

    def __init__(self, cached: Workbook, formulas: Workbook) -> None:
        self.workbook = formulas
        self.sheets = {ws.title.casefold(): ws for ws in worksheets(formulas)}
        self.cached_sheets = {ws.title.casefold(): ws for ws in worksheets(cached)}
        self.memo: dict[CellKey, Scalar | UncalculableError] = {}
        self.trees: dict[CellKey, Node] = {}
        self.active: set[CellKey] = set()
        self.work = 0
        self.depth = 0
        self.name_depth = 0
        self.array_mode = False  # whether the lazy function being called is in an array context
        self.scopes: list[dict[str, Value]] = []
        self.here = Position(next(iter(self.sheets.values())), 1, 1)

    # -- cells -----------------------------------------------------------------------------

    def calculate(self, sheet: Worksheet, row: int, col: int) -> Scalar:
        """The value of a cell, calculating its formula if Excel stored no result.

        The formula cells a cell refers to are calculated first, from an explicit stack, so
        a chain of any length (a running balance) needs no deep recursion. References the
        scan cannot see (OFFSET, INDEX) fall back to recursion, which is depth-limited.
        """
        stack = [(sheet, row, col)]
        queued = {(sheet.title, row, col)}
        while True:
            current = stack[-1]
            waiting = [d for d in self.unfinished_dependencies(*current) if _key(d) not in queued]
            if waiting:
                stack.extend(reversed(waiting))
                queued.update(_key(d) for d in waiting)
                continue
            try:
                value = self.cell_value(*current)
            except TooDeepError as deep:
                stack.append(deep.target)
                queued.add(_key(deep.target))
                continue
            except UncalculableError:
                # A failed dependency is remembered; the cells that use it will report it.
                if len(stack) == 1:
                    raise
                value = None
            except (ArithmeticError, ValueError, TypeError, RecursionError) as error:
                logger.debug("calculation failed: %r", error)
                failure = UncalculableError("not supported")
                self.memo[_key(current)] = failure
                if len(stack) == 1:
                    raise failure from None
                value = None
            stack.pop()
            queued.discard(_key(current))
            if not stack:
                return value

    def unfinished_dependencies(
        self, sheet: Worksheet, row: int, col: int
    ) -> list[tuple[Worksheet, int, int]]:
        """Formula cells that the formula at this position names and that are not done yet."""
        cell = sheet._cells.get((row, col))
        if cell is None or cell.data_type != "f" or isinstance(cell.value, ArrayFormula):
            return []
        key = (sheet.title, row, col)
        if key in self.memo or self.cached_value(sheet, row, col) is not None:
            return []
        try:
            tree = self.tree(cell)
            references = list(self._references(tree, sheet, 0))
        except (UncalculableError, FormulaError):
            return []
        found = []
        for target, top, left, bottom, right in references:
            self.charge(1)
            for position, other in _stored_in(target, top, left, bottom, right):
                if other.data_type == "f" and self._is_unfinished(target, *position):
                    found.append((target, *position))
        return found

    def _is_unfinished(self, sheet: Worksheet, row: int, col: int) -> bool:
        return (sheet.title, row, col) not in self.memo and (
            self.cached_value(sheet, row, col) is None
        )

    def _references(
        self, node: Node, home: Worksheet, depth: int
    ) -> Iterator[tuple[Worksheet, int, int, int, int]]:
        """The rectangles a syntax tree refers to, as (sheet, top, left, bottom, right)."""
        match node:
            case Ref():
                for target in self._sheets_for(node, home):
                    yield (
                        target,
                        node.top or 1,
                        node.left or 1,
                        node.bottom or target.max_row,
                        node.right or target.max_column,
                    )
            case Name() if depth < MAX_NAME_DEPTH:
                scope = self.sheets.get((node.sheet or "").casefold(), home)
                try:
                    inner = parse(self._lookup_name(scope, node.name))
                except UncalculableError:
                    return
                yield from self._references(inner, home, depth + 1)
            case Unary(_, operand) | Percent(operand):
                yield from self._references(operand, home, depth)
            case Binary(_, left, right):
                yield from self._references(left, home, depth)
                yield from self._references(right, home, depth)
            case Call():
                for argument in node.args:
                    yield from self._references(argument, home, depth)
            case ArrayLiteral(rows):
                for row in rows:
                    for item in row:
                        yield from self._references(item, home, depth)

    def _sheets_for(self, ref: Ref, home: Worksheet) -> list[Worksheet]:
        saved = self.here
        self.here = Position(home, saved.row, saved.col)
        try:
            return self.sheets_of(ref)
        except FormulaError:
            return []
        finally:
            self.here = saved

    def tree(self, cell: Any) -> Node:
        key = (cell.parent.title, cell.row, cell.column)
        if key not in self.trees:
            self.trees[key] = parse(str(cell.value))
        return self.trees[key]

    def cell_value(self, sheet: Worksheet, row: int, col: int) -> Scalar:
        cell = sheet._cells.get((row, col))
        if cell is None:
            return None
        if cell.data_type != "f":
            return constant(cell)
        key = (sheet.title, row, col)
        if key in self.memo:
            remembered = self.memo[key]
            if isinstance(remembered, UncalculableError):
                raise remembered
            return remembered
        stored = self.cached_value(sheet, row, col)
        if stored is not None:
            self.memo[key] = stored
            return stored
        if key in self.active:
            raise UncalculableError("circular reference")
        if self.depth >= MAX_DEPTH:
            raise TooDeepError(sheet, row, col)
        self.charge(1)
        outer, outer_scopes = self.here, self.scopes
        self.here, self.scopes = Position(sheet, row, col), []
        self.active.add(key)
        self.depth += 1
        try:
            result = self.evaluate_cell(cell)
        except UncalculableError as error:
            self.memo[key] = error
            raise
        finally:
            self.depth -= 1
            self.active.discard(key)
            self.here, self.scopes = outer, outer_scopes
        self.memo[key] = result
        return result

    def cached_value(self, sheet: Worksheet, row: int, col: int) -> Scalar:
        cached_sheet = self.cached_sheets.get(sheet.title.casefold())
        cell = cached_sheet._cells.get((row, col)) if cached_sheet else None
        return None if cell is None else constant(cell)

    def evaluate_cell(self, cell: Any) -> Scalar:
        if isinstance(cell.value, ArrayFormula):
            text = str(cell.value.text)
            return self.spill(cell.parent, cell.row, cell.column, text).rows[0][0]
        node = self.tree(cell)
        shown = self.eval(node, array=False)
        if isinstance(shown, RefGrid):
            shown = self.intersect(shown)
        value: Scalar = shown.rows[0][0] if isinstance(shown, Grid) else shown
        if isinstance(node, Binary) and not isinstance(value, ExcelError):
            left, right = self.final_operands(node)
            value = snap(node.op, left, right, value)
        return 0.0 if value is None else value

    def final_operands(self, node: Binary) -> tuple[Scalar, Scalar]:
        left = self.intersect(self.eval(node.left, False))
        right = self.intersect(self.eval(node.right, False))
        return (
            left if not isinstance(left, Grid) else None,
            right if not isinstance(right, Grid) else None,
        )

    def spill(self, sheet: Worksheet, row: int, col: int, formula: str) -> Grid:
        """The array that a formula entered as a dynamic array formula in this cell produces."""
        tree = parse(formula)
        outer = self.here
        self.here = Position(sheet, row, col)
        try:
            for target, top, left, bottom, right in list(self._references(tree, sheet, 0)):
                self.charge(1)
                for position, other in _stored_in(target, top, left, bottom, right):
                    if other.data_type == "f" and self._is_unfinished(target, *position):
                        self._calculate_quietly(target, *position)
            try:
                value = self.eval(tree, array=True)
            except FormulaError as error:
                value = error.error
        finally:
            self.here = outer
        rows = value.rows if isinstance(value, Grid) else [[value]]
        return Grid([[0 if item is None else item for item in line] for line in rows])

    def _calculate_quietly(self, sheet: Worksheet, row: int, col: int) -> None:
        """Calculate a cell a formula uses. A failure is remembered, and the formula reports it
        if it needs the cell."""
        with suppress(UncalculableError):
            self.calculate(sheet, row, col)

    def charge(self, units: int) -> None:
        self.work += units
        if self.work > MAX_WORK:
            raise UncalculableError("too much to calculate")

    # -- references ------------------------------------------------------------------------

    def find_sheet(self, name: str) -> Worksheet:
        sheet = self.sheets.get(name.casefold())
        if sheet is None:
            raise FormulaError(REF)
        return sheet

    def sheets_of(self, ref: Ref) -> list[Worksheet]:
        if not ref.sheets:
            return [self.here.sheet]
        name = ref.sheets[0]
        if name.casefold() not in self.sheets and ":" in name:
            first, last = (self.find_sheet(part) for part in name.split(":", 1))
            order = list(self.sheets.values())
            low, high = sorted((order.index(first), order.index(last)))
            return order[low : high + 1]
        return [self.find_sheet(name)]

    def reference(self, ref: Ref) -> RefGrid:
        sheets = self.sheets_of(ref)
        rows: list[list[Value]] = []
        for sheet in sheets:
            top = ref.top or 1
            left = ref.left or 1
            bottom = ref.bottom or max(sheet.max_row, top)
            right = ref.right or max(sheet.max_column, left)
            self.charge((bottom - top + 1) * (right - left + 1))
            for row in range(top, bottom + 1):
                rows.append([self.cell_value(sheet, row, col) for col in range(left, right + 1)])
        return RefGrid(
            rows,
            ref.top or 1,
            ref.left or 1,
            len(sheets),
            clipped_rows=ref.top is None or ref.bottom is None,
            clipped_cols=ref.left is None or ref.right is None,
        )

    def intersect(self, value: Value) -> Value:
        """Excel's implicit intersection of a range with the formula's own cell."""
        if not isinstance(value, RefGrid):
            return value
        if value.height == 1 and value.width == 1:
            return value.rows[0][0]
        if value.sheets > 1:
            raise UncalculableError("3-D reference in a single-value position")
        row = self.here.row - value.top
        col = self.here.col - value.left
        if value.width == 1:
            col = 0
        if value.height == 1:
            row = 0
        if 0 <= row < value.height and 0 <= col < value.width:
            return value.rows[row][col]
        return VALUE

    def name(self, node: Name, array: bool) -> Value:
        for scope in reversed(self.scopes):
            if node.name.casefold() in scope:
                return scope[node.name.casefold()]
        if self.name_depth >= MAX_NAME_DEPTH:
            raise UncalculableError("names nested too deeply")
        scope = self.find_sheet(node.sheet) if node.sheet else self.here.sheet
        defined = self._lookup_name(scope, node.name)
        tree = parse(defined)
        if _has_relative_reference(tree):
            raise UncalculableError(f"name {node.name} uses relative references")
        self.name_depth += 1
        try:
            return self.eval(tree, array)
        finally:
            self.name_depth -= 1

    def _lookup_name(self, scope: Worksheet, name: str) -> str:
        for names in (scope.defined_names, self.workbook.defined_names):
            for key, defined in names.items():
                if key.casefold() == name.casefold() and defined.attr_text:
                    return "=" + defined.attr_text
        raise UncalculableError(f"name {name}")

    # -- expressions -----------------------------------------------------------------------

    def eval(self, node: Node, array: bool = False) -> Value:
        match node:
            case Literal(value):
                return value
            case ErrorLiteral(code):
                if code not in ERRORS:
                    raise UncalculableError(f"error value {code}")
                return ERRORS[code]
            case Ref():
                try:
                    return self.reference(node)
                except FormulaError as error:
                    return error.error
            case Name():
                try:
                    return self.name(node, array)
                except FormulaError as error:
                    return error.error
            case Unary(op, operand):
                return elementwise(lambda v: apply_unary(op, v), self.operand(operand, array))
            case Percent(operand):
                return elementwise(apply_percent, self.operand(operand, array))
            case Binary(op, left, right):
                return elementwise(
                    lambda a, b: apply_binary(op, a, b),
                    self.operand(left, array),
                    self.operand(right, array),
                )
            case Call():
                return self.call(node, array)
            case ArrayLiteral(rows):
                return Grid([[self.scalar(item) for item in row] for row in rows])
        raise UncalculableError("unsupported syntax")

    def operand(self, node: Node, array: bool) -> Value:
        value = self.eval(node, array)
        return value if array else self.intersect(value)

    def scalar(self, node: Node) -> Scalar:
        """Evaluate a node where one value is expected."""
        value = self.operand(node, array=False)
        if isinstance(value, Grid):
            raise UncalculableError("array where a single value is expected")
        return value

    def call(self, node: Call, array: bool) -> Value:
        spec = FUNCTIONS.get(node.name)
        if spec is None:
            raise UncalculableError(node.name)
        if node.name in FUTURE_FUNCTIONS and not node.prefixed:
            raise UncalculableError(f"{node.name} needs the {FUTURE_FUNCTIONS[node.name]} prefix")
        if not spec.min_args <= len(node.args) <= spec.max_args:
            raise UncalculableError(f"{node.name}: wrong number of arguments")
        self.charge(1)
        try:
            result = self.dispatch(spec, node.args, array)
        except FormulaError as error:
            return error.error
        if isinstance(result, float) and not math.isfinite(result):
            return NUM
        return result

    def dispatch(self, spec: Function, args: tuple[Node, ...], array: bool) -> Value:
        if spec.kind == "lazy":
            outer, self.array_mode = self.array_mode, array
            try:
                return spec.call(self, *args)
            finally:
                self.array_mode = outer
        values = [self.eval(arg, array or spec.array_at(index)) for index, arg in enumerate(args)]
        if spec.kind in ("scalar", "check"):
            return self.call_scalar(spec, values, array)
        if spec.kind == "eager":
            for value in values:
                if isinstance(value, ExcelError):
                    return value
        return spec.call(*values)

    def call_scalar(self, spec: Function, values: list[Value], array: bool) -> Value:
        if not array:
            values = [self.intersect(value) for value in values]
            if any(isinstance(value, Grid) for value in values):
                raise UncalculableError("array where a single value is expected")

        def apply(*scalars: Scalar) -> Scalar:
            if spec.kind == "scalar":
                for value in scalars:
                    if isinstance(value, ExcelError):
                        return value
            try:
                return spec.call(*scalars)
            except FormulaError as error:
                return error.error

        return elementwise(apply, *values)


def _key(position: tuple[Worksheet, int, int]) -> CellKey:
    return (position[0].title, position[1], position[2])


def _stored_in(
    sheet: Worksheet, top: int, left: int, bottom: int, right: int
) -> Iterator[tuple[tuple[int, int], Any]]:
    """The stored cells inside a rectangle, whichever is fewer to walk: it or the sheet."""
    if (bottom - top + 1) * (right - left + 1) <= len(sheet._cells):
        for row in range(top, bottom + 1):
            for col in range(left, right + 1):
                if (cell := sheet._cells.get((row, col))) is not None:
                    yield (row, col), cell
    else:
        for (row, col), cell in sheet._cells.items():
            if top <= row <= bottom and left <= col <= right:
                yield (row, col), cell


def constant(cell: Any) -> Scalar:
    value = cell.value
    if cell.data_type == "e":
        return ERRORS.get(str(value), VALUE)
    match value:
        case dt.datetime():
            return (
                date_to_serial(value.date())
                + (value - dt.datetime.combine(value.date(), dt.time())).total_seconds() / 86400
            )
        case dt.date():
            return date_to_serial(value)
        case dt.time():
            return (value.hour * 3600 + value.minute * 60 + value.second) / 86400
        case dt.timedelta():
            return value.total_seconds() / 86400
        case Decimal():
            return float(value)
        case bool() | int() | float() | str() | None:
            return value
    raise UncalculableError("unsupported cell content")


def _has_relative_reference(node: Node) -> bool:
    match node:
        case Ref():
            return node.relative
        case Unary(_, operand) | Percent(operand):
            return _has_relative_reference(operand)
        case Binary(_, left, right):
            return _has_relative_reference(left) or _has_relative_reference(right)
        case Call():
            return any(_has_relative_reference(arg) for arg in node.args)
        case _:
            return False
