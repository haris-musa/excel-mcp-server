"""Database functions: DSUM, DCOUNT and the rest, with Excel's criteria ranges."""

import math
import re
from collections.abc import Callable
from typing import TYPE_CHECKING

from openpyxl.formula.tokenizer import Token, Tokenizer
from openpyxl.formula.translate import Translator
from openpyxl.utils import column_index_from_string

from excel_mcp.calc.criteria import Test, criterion
from excel_mcp.calc.functions.math import naive_sum
from excel_mcp.calc.parser import Literal, Node, Ref
from excel_mcp.calc.registry import function
from excel_mcp.calc.values import (
    DIV0,
    NUM,
    VALUE,
    ExcelError,
    FormulaError,
    RefGrid,
    Scalar,
    UncalculableError,
    is_number,
    parse_number,
    scalar,
)

if TYPE_CHECKING:
    from excel_mcp.calc.engine import Engine

_OPERATORS = ("<", ">", "=")
_CELL = re.compile(
    r"(?:(?P<sheet>'(?:[^']|'')+'|[^'!]+)!)?\$?(?P<col>[A-Za-z]{1,3})\$?(?P<row>\d+)"
)


def _range(engine: "Engine", node: Node, rows: int) -> RefGrid:
    value = engine.eval(node)
    if not isinstance(value, RefGrid) or value.height < rows:
        raise FormulaError(VALUE)
    return value


def _test(raw: Scalar) -> Test:
    """A criterion as the database functions read it: plain text matches a start."""
    if (
        isinstance(raw, str)
        and not raw.startswith(_OPERATORS)
        and parse_number(raw) is None
        and raw.upper() not in ("TRUE", "FALSE")
        and not raw.startswith("#")
    ):
        return criterion(raw + "*")
    return criterion(raw)


def _blank(cell: Scalar) -> bool:
    return cell is None or (isinstance(cell, str) and cell != "" and not cell.strip())


class Database:
    """The records of a database range and the criteria range that selects some of them."""

    def __init__(self, engine: "Engine", database: RefGrid, criteria: RefGrid) -> None:
        if criteria.sheet is None or database.sheet is None:
            raise UncalculableError("database over several sheets")
        self.engine, self.source, self.criteria = engine, database, criteria
        self.headers = [h.casefold() if isinstance(h, str) else "" for h in database.rows[0]]
        self.records = database.rows[1:]

    def column(self, name: str) -> int | None:
        wanted = name.casefold()
        return self.headers.index(wanted) if wanted in self.headers else None

    def selected(self) -> list[list[Scalar]]:
        """The records that satisfy every cell of some row of the criteria."""
        labels = [h if isinstance(h, str) else "" for h in self.criteria.rows[0]]
        targets = [self.column(label) if label else None for label in labels]
        alternatives = [
            [(c, targets[c], _test(cell)) for c, cell in enumerate(row) if not _blank(cell)]  # pyright: ignore[reportArgumentType]
            for row in self.criteria.rows[1:]
        ]
        return [
            record  # pyright: ignore[reportReturnType]
            for offset, record in enumerate(self.records)
            if any(
                all(
                    test(record[target]) if target is not None else self._computed(r + 1, c, offset)
                    for c, target, test in row
                )
                for r, row in enumerate(alternatives)
            )
        ]

    def _computed(self, row: int, col: int, offset: int) -> bool:
        """A criteria cell under a label that is no field: a formula is evaluated as if it
        referred to the record ``offset`` rows below the first; a number or TRUE is its own
        result; text, an error or FALSE does not select."""
        sheet = self.engine.sheets[(self.criteria.sheet or "").casefold()]
        top, left = self.criteria.top + row, self.criteria.left + col
        cell = sheet._cells[(top, left)]
        result: Scalar = cell.value  # pyright: ignore[reportAssignmentType]
        if cell.data_type == "f":
            formula = self._shifted(str(cell.value), sheet.title, offset)
            result = self.engine.evaluate_formula(sheet, top, left, formula)
        return isinstance(result, bool | int | float) and bool(result)

    def _shifted(self, formula: str, home: str, offset: int) -> str:
        """The formula with the references to cells of the database moved down ``offset`` rows."""
        tokens = Tokenizer(formula)
        for token in tokens.items:
            if token.type == Token.OPERAND and token.subtype == Token.RANGE:
                match = _CELL.fullmatch(token.value)
                if match and self._in_records(match, home):
                    origin, target = "A1", f"A{1 + offset}"
                    token.value = Translator("=" + token.value, origin).translate_formula(target)[
                        1:
                    ]
        return tokens.render()

    def _in_records(self, match: re.Match[str], home: str) -> bool:
        name = (match["sheet"] or home).strip("'").replace("''", "'")
        column = column_index_from_string(match["col"].upper())
        source = self.source
        return (
            name.casefold() == (source.sheet or "").casefold()
            and source.top < int(match["row"]) < source.top + source.height
            and source.left <= column < source.left + source.width
        )


def _field_index(database: Database, field: Scalar) -> int:
    if isinstance(field, str):
        index = database.column(field)
    elif field is None or isinstance(field, ExcelError):
        index = None
    else:
        index = int(field) - 1  # a number, or TRUE as 1
        index = index if 0 <= index < len(database.headers) else None
    if index is None:
        raise FormulaError(VALUE)
    return index


def _field(engine: "Engine", node: Node) -> Scalar:
    """The field argument; an error held in a cell is #VALUE!, one given directly is the result."""
    value = scalar(engine.eval(node))
    if isinstance(value, ExcelError):
        raise FormulaError(VALUE if isinstance(node, Ref) else value)
    return value


def _open(engine: "Engine", database: Node, criteria: Node) -> Database:
    return Database(engine, _range(engine, database, 2), _range(engine, criteria, 2))


def _column(engine: "Engine", database: Node, field: Node, criteria: Node) -> list[Scalar]:
    source = _open(engine, database, criteria)
    index = _field_index(source, _field(engine, field))
    return [record[index] for record in source.selected()]


def _numbers(values: list[Scalar]) -> list[float]:
    for value in values:
        if isinstance(value, ExcelError):
            raise FormulaError(value)
    return [float(v) for v in values if is_number(v)]  # pyright: ignore[reportArgumentType]


def _aggregate(name: str, reducer: Callable[[list[float]], float]) -> None:
    def call(engine: "Engine", database: Node, field: Node, criteria: Node) -> float:
        return reducer(_numbers(_column(engine, database, field, criteria)))

    function(name, kind="lazy")(call)


def _variance(sample: bool) -> Callable[[list[float]], float]:
    def variance(values: list[float]) -> float:
        n = len(values)
        if n < (2 if sample else 1):
            raise FormulaError(DIV0)
        mean = naive_sum(values) / n
        return naive_sum((v - mean) ** 2 for v in values) / (n - 1 if sample else n)

    return variance


def _average(values: list[float]) -> float:
    if not values:
        raise FormulaError(DIV0)
    return naive_sum(values) / len(values)


_aggregate("DSUM", naive_sum)
_aggregate("DAVERAGE", _average)
_aggregate("DMAX", lambda v: max(v, default=0.0))
_aggregate("DMIN", lambda v: min(v, default=0.0))
_aggregate("DPRODUCT", lambda v: math.prod(v) if v else 0.0)
_aggregate("DVAR", _variance(sample=True))
_aggregate("DVARP", _variance(sample=False))
_aggregate("DSTDEV", lambda v: math.sqrt(_variance(sample=True)(v)))
_aggregate("DSTDEVP", lambda v: math.sqrt(_variance(sample=False)(v)))


def _counted(engine: "Engine", database: Node, field: Node, criteria: Node, numeric: bool) -> float:
    source = _open(engine, database, criteria)
    if isinstance(field, Literal) and field.value is None:  # the field left out: count records
        return float(len(source.selected()))
    values = [r[_field_index(source, _field(engine, field))] for r in source.selected()]
    return float(len(_numbers(values)) if numeric else sum(v is not None for v in values))


@function("DCOUNT", kind="lazy")
def dcount(engine: "Engine", database: Node, field: Node, criteria: Node) -> float:
    return _counted(engine, database, field, criteria, numeric=True)


@function("DCOUNTA", kind="lazy")
def dcounta(engine: "Engine", database: Node, field: Node, criteria: Node) -> float:
    return _counted(engine, database, field, criteria, numeric=False)


@function("DGET", kind="lazy")
def dget(engine: "Engine", database: Node, field: Node, criteria: Node) -> Scalar:
    found = _column(engine, database, field, criteria)
    if len(found) > 1:
        raise FormulaError(NUM)
    if not found or found[0] is None:
        raise FormulaError(VALUE)
    return found[0]
