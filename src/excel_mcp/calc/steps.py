"""A formula's sub-expressions and their values in the order Excel's Evaluate Formula shows them."""

import re
from dataclasses import dataclass

from openpyxl.utils import get_column_letter
from openpyxl.worksheet.formula import ArrayFormula
from openpyxl.worksheet.worksheet import Worksheet

from excel_mcp.calc.dependencies import formula_text, formula_tree
from excel_mcp.calc.engine import Engine, Position
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
    TableRef,
    Unary,
    children,
    parse,
)
from excel_mcp.calc.registry import FUNCTIONS
from excel_mcp.calc.values import (
    NA,
    ExcelError,
    FormulaError,
    Grid,
    Scalar,
    UncalculableError,
    Value,
    number_text,
    to_bool,
    to_int,
)
from excel_mcp.refs import cell_name

_PRECEDENCE = {"=": 1, "<>": 1, "<": 1, ">": 1, "<=": 1, ">=": 1, "&": 2, "+": 3, "-": 3}
_PRECEDENCE |= {"*": 4, "/": 4, "^": 5}
_ATOMIC = frozenset(
    ["LET", "LAMBDA", "MAP", "REDUCE", "SCAN", "BYROW", "BYCOL", "MAKEARRAY", "SWITCH"]
)
_SHORT_CIRCUITS = frozenset(["IF", "IFERROR", "IFNA", "CHOOSE", "IFS"])
_SIMPLE_SHEET = re.compile(r"\w+")
MAX_SHOWN_ITEMS = 30


@dataclass
class Steps:
    """The recorded steps, the first ``limit`` and the last, and how many were left out."""

    shown: list[tuple[str, Value]]
    skipped: int


def evaluate(engine: Engine, sheet: Worksheet, row: int, col: int, limit: int) -> Steps:
    """Every sub-expression of the formula at this position with its value, innermost first.

    Only the branches Excel would take are followed. Raises ``UncalculableError`` when the
    calculator cannot evaluate the formula.
    """
    cell = sheet._cells[(row, col)]
    tree = formula_tree(engine, cell)
    array = isinstance(cell.value, ArrayFormula)
    outer, engine.here = engine.here, Position(sheet, row, col)
    try:
        engine.scheduler.calculate_precedents(tree, sheet)
        walk = _Walk(engine, limit)
        walk.visit(tree, parse(formula_text(cell)), array, final=True)
        return Steps(walk.shown, walk.skipped)
    finally:
        engine.here = outer


class _Walk:
    """Visits a formula's tree (table references replaced by their cells) together with the
    tree as written, which is what the steps show."""

    def __init__(self, engine: Engine, limit: int) -> None:
        self.engine = engine
        self.limit = limit
        self.shown: list[tuple[str, Value]] = []
        self.skipped = 0

    def value(self, node: Node, array: bool) -> Value:
        self.engine.work = 0
        return self.engine.eval(node, array)

    def visit(self, node: Node, written: Node, array: bool, *, final: bool = False) -> None:
        """Record the steps inside the node, then the node. The whole formula is always kept."""
        if isinstance(node, Literal | ErrorLiteral) and not final:
            return
        mark = len(self.shown), self.skipped
        try:
            self._descend(node, written, array)
        except UncalculableError:
            # Something inside cannot be evaluated alone (a LET variable): show the node whole.
            del self.shown[mark[0] :]
            self.skipped = mark[1]
        value = self.value(node, array)
        if len(self.shown) < self.limit - 1 or final:
            self.shown.append((render(written), value))
        else:
            self.skipped += 1

    def _descend(self, node: Node, written: Node, array: bool) -> None:
        pairs = list(zip(children(node), children(written), strict=True))
        if not isinstance(node, Call):
            for child, shown_child in pairs:
                self.visit(child, shown_child, array)
        elif node.name in _SHORT_CIRCUITS and pairs:
            self._choose(node.name, pairs, array)
        elif node.name not in _ATOMIC:
            spec = FUNCTIONS.get(node.name)
            for index, (child, shown_child) in enumerate(pairs):
                self.visit(child, shown_child, array or (spec is not None and spec.array_at(index)))

    def _choose(self, name: str, pairs: list[tuple[Node, Node]], array: bool) -> None:
        if name == "IFS":
            for (condition, shown_condition), (result, shown_result) in zip(
                pairs[::2], pairs[1::2], strict=False
            ):
                self.visit(condition, shown_condition, array)
                if self._truth(condition, array):
                    self.visit(result, shown_result, array)
                    return
            return
        self.visit(*pairs[0], array)
        first = self.value(pairs[0][0], array)
        if name == "IF":
            pick = 1 if self._truth(pairs[0][0], array) else 2
        elif name == "CHOOSE":
            pick = to_int(first) if not isinstance(first, ExcelError | Grid) else 0
        else:
            failed = isinstance(first, ExcelError) and (name == "IFERROR" or first == NA)
            pick = 1 if failed else 0
        if 0 < pick < len(pairs):
            self.visit(*pairs[pick], array)

    def _truth(self, node: Node, array: bool) -> bool:
        try:
            return to_bool(self.value(node, array))
        except FormulaError:  # not a boolean: the formula's own step shows the error
            return False


def render(node: Node) -> str:
    """The formula text of a node."""
    match node:
        case Literal(value):
            return _literal(value)
        case ErrorLiteral(code):
            return code
        case Ref():
            return _reference(node)
        case TableRef(ref):
            return ref.long_form(ref.table or "")
        case Name(sheet, name):
            return f"{sheet_text(sheet)}!{name}" if sheet else name
        case Unary(op, operand):
            return op + _wrapped(operand, 6)
        case Percent(operand):
            return _wrapped(operand, 6) + "%"
        case Binary(op, left, right):
            level = _PRECEDENCE[op]
            return f"{_wrapped(left, level)}{op}{_wrapped(right, level + 1)}"
        case Call("ANCHORARRAY", (Ref() as anchor,), _):
            return render(anchor) + "#"
        case Call(name, args, _):
            return f"{name}({','.join(render(a) for a in args)})"
        case ArrayLiteral(rows):
            return "{" + ";".join(",".join(render(i) for i in row) for row in rows) + "}"
    raise UncalculableError("unsupported syntax")


def _wrapped(node: Node, level: int) -> str:
    inner = render(node)
    if isinstance(node, Binary) and _PRECEDENCE[node.op] < level:
        return f"({inner})"
    if isinstance(node, Unary) and level > 6:
        return f"({inner})"
    return inner


def _literal(value: float | str | bool | None) -> str:
    if value is None:
        return ""
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, str):
        return '"' + value.replace('"', '""') + '"'
    return number_text(value)


def sheet_text(name: str) -> str:
    return name if _SIMPLE_SHEET.fullmatch(name) and not name[0].isdigit() else f"'{name}'"


def _reference(ref: Ref) -> str:
    prefix = f"{sheet_text(ref.sheets[0])}!" if ref.sheets else ""
    if ref.top is None or ref.bottom is None:
        return f"{prefix}{get_column_letter(ref.left or 1)}:{get_column_letter(ref.right or 1)}"
    if ref.left is None or ref.right is None:
        return f"{prefix}{ref.top}:{ref.bottom}"
    start, end = cell_name(ref.top, ref.left), cell_name(ref.bottom, ref.right)
    return prefix + (start if start == end else f"{start}:{end}")


def shown(value: Value) -> str:
    """A step's value as Excel's Evaluate Formula writes it: text quoted, errors by code,
    arrays in braces."""
    if isinstance(value, Grid):
        if value.height == 1 and value.width == 1:
            return scalar_text(value.rows[0][0])
        return _array(value)
    return scalar_text(value)


def _array(grid: Grid) -> str:
    rows = [[scalar_text(value) for value in row] for row in grid.rows]
    if grid.height * grid.width > MAX_SHOWN_ITEMS:
        kept = [item for row in rows for item in row][:MAX_SHOWN_ITEMS]
        return f"{{{','.join(kept)},...}} {grid.height}x{grid.width}"
    return "{" + ";".join(",".join(row) for row in rows) + "}"


def scalar_text(value: Scalar) -> str:
    return value.code if isinstance(value, ExcelError) else _literal(value)
