"""Evaluates formulas against the cells of a workbook, the way Excel would.

Anything that cannot be reproduced exactly raises `UncalculableError`; the engine never guesses.
Formulas are evaluated as Excel evaluates ordinary (not array-entered) formulas: a range
where a single value is expected is reduced to the formula cell's own row or column.
"""

import logging
import math
from dataclasses import dataclass, replace
from typing import Any

from openpyxl import Workbook
from openpyxl.worksheet.formula import ArrayFormula
from openpyxl.worksheet.worksheet import Worksheet

from excel_mcp.calc import functions as functions
from excel_mcp.calc.cells import CellKey, constant
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
from excel_mcp.calc.scheduler import MAX_NAME_DEPTH, Scheduler, TooDeepError
from excel_mcp.calc.tablerefs import bind
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
)
from excel_mcp.lazy_workbook import extent
from excel_mcp.refs import parse_range
from excel_mcp.workspace import worksheets
from excel_mcp.xlfn import FUTURE_FUNCTIONS

logger = logging.getLogger(__name__)

MAX_WORK = 1_000_000
TEXT_PER_UNIT = 16
MAX_DEPTH = 30


@dataclass
class Position:
    sheet: Worksheet
    row: int
    col: int


class Engine:
    """Calculates formula cells of ``formulas``; ``cached`` holds Excel's stored results."""

    def __init__(self, cached: Workbook, formulas: Workbook, filename: str | None = None) -> None:
        self.workbook = formulas
        self.filename = filename
        self.sheets = {ws.title.casefold(): ws for ws in worksheets(formulas)}
        self.cached_sheets = {ws.title.casefold(): ws for ws in worksheets(cached)}
        self.memo: dict[CellKey, Scalar | UncalculableError] = {}
        self.trees: dict[CellKey, Node] = {}
        self.spills: dict[CellKey, Grid] = {}
        self.active: set[CellKey] = set()
        self.work = 0
        self.depth = 0
        self.name_depth = 0
        self.array_mode = False  # whether the lazy function being called is in an array context
        self.scopes: list[dict[str, Value]] = []
        self.here = Position(next(iter(self.sheets.values())), 1, 1)
        self.scheduler = Scheduler(self)

    # -- cells -----------------------------------------------------------------------------

    def calculate(self, sheet: Worksheet, row: int, col: int) -> Scalar:
        """The value of a cell, calculating its formula if Excel stored no result."""
        return self.scheduler.calculate(sheet, row, col)

    def sheets_in(self, ref: Ref, home: Worksheet) -> list[Worksheet]:
        """The sheets a reference names, read as if the formula sat on ``home``."""
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
            tree = parse(str(cell.value))
            self.trees[key] = bind(tree, self.workbook, cell.parent, cell.row, cell.column)
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
            return self.spill_of(cell.parent, cell.row, cell.column, text).rows[0][0]
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

    def spill_of(self, sheet: Worksheet, row: int, col: int, formula: str) -> Grid:
        key = (sheet.title, row, col)
        if key not in self.spills:
            self.spills[key] = self.spill(sheet, row, col, formula)
        return self.spills[key]

    def anchored(self, ref: Ref) -> Value:
        """What the spill reference ``A1#`` stands for: the range the formula in A1 filled.

        Where Excel stored the result, that is the stored array range; otherwise it is what
        the calculator gets for the formula.
        """
        if ref.top is None or ref.left is None or (ref.top, ref.left) != (ref.bottom, ref.right):
            raise UncalculableError("spill reference to something other than one cell")
        sheet = self.sheets_of(ref)[0]
        cell = sheet._cells.get((ref.top, ref.left))
        if cell is None or cell.data_type != "f":
            raise FormulaError(REF)
        if not isinstance(cell.value, ArrayFormula):
            return self.reference(ref)  # Excel: a formula that spills nothing fills its own cell
        if self.cached_value(sheet, ref.top, ref.left) is None:
            return self.spill_of(sheet, ref.top, ref.left, str(cell.value.text))
        area = parse_range(cell.value.ref)
        return self.reference(
            replace(
                ref, top=area.min_row, left=area.min_col, bottom=area.max_row, right=area.max_col
            )
        )

    def spill(self, sheet: Worksheet, row: int, col: int, formula: str) -> Grid:
        """The array that a formula entered as a dynamic array formula in this cell produces."""
        tree = bind(parse(formula), self.workbook, sheet, row, col)
        outer = self.here
        self.here = Position(sheet, row, col)
        try:
            self.scheduler.calculate_precedents(tree, sheet)
            try:
                value = self.eval(tree, array=True)
            except FormulaError as error:
                value = error.error
        finally:
            self.here = outer
        rows = value.rows if isinstance(value, Grid) else [[value]]
        return Grid([[0 if item is None else item for item in line] for line in rows])

    def sized(self, value: Scalar) -> Scalar:
        """Count produced text against the work limit, so arrays of long text stop."""
        if isinstance(value, str):
            self.charge(len(value) // TEXT_PER_UNIT)
        return value

    def charge(self, units: int) -> None:
        # A request that is too big is refused without using up what is left for other cells.
        if self.work + units > MAX_WORK:
            raise UncalculableError("too much to calculate")
        self.work += units

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
            last_row, last_column = extent(sheet)
            bottom = ref.bottom or max(last_row, top)
            right = ref.right or max(last_column, left)
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
        # A whole row or column stops at the last used cell; beyond it the cells are blank.
        rows_in = 0 <= row < value.height
        cols_in = 0 <= col < value.width
        past_rows = value.clipped_rows and row >= value.height and cols_in
        past_cols = value.clipped_cols and col >= value.width and rows_in
        if rows_in and cols_in:
            return value.rows[row][col]
        if past_rows or past_cols:
            return None
        return VALUE

    def name(self, node: Name, array: bool) -> Value:
        for scope in reversed(self.scopes):
            if node.name.casefold() in scope:
                return scope[node.name.casefold()]
        if self.name_depth >= MAX_NAME_DEPTH:
            raise UncalculableError("names nested too deeply")
        scope = self.find_sheet(node.sheet) if node.sheet else self.here.sheet
        defined = self.lookup_name(scope, node.name)
        tree = bind(parse(defined), self.workbook, self.here.sheet, self.here.row, self.here.col)
        if _has_relative_reference(tree):
            raise UncalculableError(f"name {node.name} uses relative references")
        self.name_depth += 1
        try:
            return self.eval(tree, array)
        finally:
            self.name_depth -= 1

    def is_defined(self, node: Name) -> bool:
        scope = self.find_sheet(node.sheet) if node.sheet else self.here.sheet
        try:
            self.lookup_name(scope, node.name)
        except UncalculableError:
            return False
        return True

    def lookup_name(self, scope: Worksheet, name: str) -> str:
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
                    lambda a, b: self.sized(apply_binary(op, a, b)),
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
        if isinstance(result, str):
            self.sized(result)
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
                return self.sized(spec.call(*scalars))
            except FormulaError as error:
                return error.error

        return elementwise(apply, *values)


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
