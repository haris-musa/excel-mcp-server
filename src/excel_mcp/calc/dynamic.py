"""Which formulas Excel stores as dynamic array formulas.

Excel saves a formula with ``t="array"`` and the dynamic array metadata when evaluating it
the modern way gives another answer than the old one, which reduced a range to one value
at the formula's own row or column. That is the case when the formula can return several
values, when an operator or a function that expects single values receives a range, and
for functions that only exist for arrays. The rules come from what Excel itself writes.
"""

from collections.abc import Callable

from excel_mcp.calc import functions as functions
from excel_mcp.calc.parser import (
    ArrayLiteral,
    Binary,
    Call,
    Literal,
    Name,
    Node,
    Percent,
    Ref,
    Unary,
    parse,
)
from excel_mcp.calc.registry import FUNCTIONS
from excel_mcp.calc.values import UncalculableError

# Functions whose result is an array, so a formula that ends in one spills.
_ARRAYS = frozenset(
    [
        "ANCHORARRAY",
        "SORT",
        "SORTBY",
        "UNIQUE",
        "FILTER",
        "SEQUENCE",
        "RANDARRAY",
        "TRANSPOSE",
        "TAKE",
        "DROP",
        "CHOOSECOLS",
        "CHOOSEROWS",
        "HSTACK",
        "VSTACK",
        "TOCOL",
        "TOROW",
        "WRAPCOLS",
        "WRAPROWS",
        "EXPAND",
        "TEXTSPLIT",
        "MMULT",
        "MINVERSE",
        "MUNIT",
        "FREQUENCY",
        "LINEST",
        "LOGEST",
        "TREND",
        "GROWTH",
        "MAP",
        "BYROW",
        "BYCOL",
        "SCAN",
        "MAKEARRAY",
    ]
)
# Functions that need array evaluation even when another function consumes their result.
_EVALUATED_AS_ARRAYS = frozenset(
    ["TRANSPOSE", "MAP", "BYROW", "BYCOL", "REDUCE", "SCAN", "MAKEARRAY"]
)
# Arguments (by position) that are single values although the function reads ranges.
_SINGLE_VALUES = {
    "LARGE": (1,),
    "SMALL": (1,),
    "COUNTIF": (1,),
    "SUMIF": (1,),
    "AVERAGEIF": (1,),
    "PERCENTILE": (1,),
    "QUARTILE": (1,),
    "INDEX": (1, 2),
    "MATCH": (0,),
    "XMATCH": (0,),
    "VLOOKUP": (0,),
    "HLOOKUP": (0,),
    "LOOKUP": (0,),
    "XLOOKUP": (0,),
}
_CHOOSING = frozenset(["IF", "IFS", "IFERROR", "IFNA", "SWITCH", "CHOOSE"])
_PASSING_THROUGH = frozenset(["N", "T"])

NameResolver = Callable[[Name], str | None]


def is_dynamic_array(formula: str, resolve: NameResolver) -> bool:
    """Whether Excel would store ``formula`` (starting with ``=``) as a dynamic array formula."""
    try:
        tree = parse(formula)
    except UncalculableError:
        return False
    analysis = _Analysis(resolve)
    return analysis.array(tree, {}) or analysis.needs_arrays


class _Analysis:
    def __init__(self, resolve: NameResolver) -> None:
        self.resolve = resolve
        self.needs_arrays = False
        self.depth = 0

    def array(self, node: Node, scope: dict[str, bool]) -> bool:
        """Whether the node can hold several values; notes where that requires array evaluation."""
        match node:
            case Ref():
                return _is_range(node)
            case ArrayLiteral(rows):
                return len(rows) > 1 or any(len(row) > 1 for row in rows)
            case Name():
                return self._name(node, scope)
            case Unary(_, operand) | Percent(operand):
                return self._operator(self.array(operand, scope))
            case Binary(_, left, right):
                return self._operator(self.array(left, scope), self.array(right, scope))
            case Call():
                return self._call(node, scope)
        return False

    def _operator(self, *operands: bool) -> bool:
        if any(operands):
            self.needs_arrays = True
        return any(operands)

    def _name(self, node: Name, scope: dict[str, bool]) -> bool:
        if node.name.casefold() in scope:
            return scope[node.name.casefold()]
        text = self.resolve(node)
        if text is None or self.depth >= 4:
            return False
        self.depth += 1
        try:
            return self.array(parse("=" + text), {})
        except UncalculableError:
            return False
        finally:
            self.depth -= 1

    def _call(self, node: Call, scope: dict[str, bool]) -> bool:
        name = node.name
        if name == "LET":
            return self._let(node, scope)
        shapes = [self.array(argument, scope) for argument in node.args]
        spec = FUNCTIONS.get(name)
        if name in _EVALUATED_AS_ARRAYS or (name in ("ROW", "COLUMN") and any(shapes)):
            self.needs_arrays = True
        if name in _ARRAYS or name in ("ROW", "COLUMN"):
            return name in _ARRAYS or any(shapes)
        if name in _SINGLE_VALUES and any(
            shapes[i] for i in _SINGLE_VALUES[name] if i < len(shapes)
        ):
            self.needs_arrays = True
            return True
        if name in _CHOOSING:
            return any(shapes)
        if name == "INDEX":
            return _index_returns_array(node)
        if name == "OFFSET":
            return any(not _is_one(a) for a in node.args[3:5]) if len(node.args) > 3 else False
        if name == "XLOOKUP":
            return _xlookup_returns_array(node, shapes)
        if spec is not None and spec.kind in ("scalar", "check") and name not in _PASSING_THROUGH:
            if any(shapes):
                self.needs_arrays = True
            return any(shapes)
        return False

    def _let(self, node: Call, scope: dict[str, bool]) -> bool:
        inner = dict(scope)
        for position in range(0, len(node.args) - 1, 2):
            variable = node.args[position]
            if isinstance(variable, Name):
                inner[variable.name.casefold()] = self.array(node.args[position + 1], inner)
        return self.array(node.args[-1], inner)


def _is_range(node: Ref) -> bool:
    return None in (node.top, node.bottom, node.left, node.right) or (
        node.top != node.bottom or node.left != node.right
    )


def _is_one(node: Node) -> bool:
    return isinstance(node, Literal) and node.value in (1, None)


def _is_zero(node: Node) -> bool:
    return isinstance(node, Literal) and node.value == 0 and node.value is not None


def _index_returns_array(node: Call) -> bool:
    """INDEX gives a whole row or column when an index is 0."""
    return any(_is_zero(argument) for argument in node.args[1:3])


def _xlookup_returns_array(node: Call, shapes: list[bool]) -> bool:
    if shapes and shapes[0]:
        return True
    if len(node.args) < 3 or not (isinstance(node.args[1], Ref) and isinstance(node.args[2], Ref)):
        return False
    searched, returned = node.args[1], node.args[2]
    if searched.left == searched.right:
        return returned.left != returned.right
    return returned.top != returned.bottom
