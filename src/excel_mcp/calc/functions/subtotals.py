"""SUBTOTAL and AGGREGATE, which skip hidden rows, nested subtotals and errors."""

from typing import TYPE_CHECKING

from excel_mcp.calc.parser import Node, Ref
from excel_mcp.calc.registry import FUNCTIONS, function
from excel_mcp.calc.values import (
    VALUE,
    ExcelError,
    FormulaError,
    Grid,
    Scalar,
    UncalculableError,
    Value,
    is_number,
    to_int,
)
from excel_mcp.lazy_workbook import extent

if TYPE_CHECKING:
    from excel_mcp.calc.engine import Engine

_SUBTOTAL = {
    1: "AVERAGE", 2: "COUNT", 3: "COUNTA", 4: "MAX", 5: "MIN", 6: "PRODUCT",
    7: "STDEV.S", 8: "STDEV.P", 9: "SUM", 10: "VAR.S", 11: "VAR.P",
}  # fmt: skip
_AGGREGATE = {**_SUBTOTAL, 12: "MEDIAN", 13: "MODE.SNGL"}
_WITH_K = {14: "LARGE", 15: "SMALL", 16: "PERCENTILE.INC", 17: "QUARTILE.INC"}
_WITH_K |= {18: "PERCENTILE.EXC", 19: "QUARTILE.EXC"}
_NESTED = ("SUBTOTAL(", "AGGREGATE(")


def _cells(
    engine: "Engine", node: Node, hidden: str, skip_nested: bool, skip_errors: bool
) -> list[Scalar]:
    """The values of a reference, leaving out what the function ignores.

    ``hidden`` is "skip" to leave out hidden rows, "keep" to include them and "unknown" when
    Excel's treatment depends on whether the rows were hidden by hand or by a filter.
    """
    if not isinstance(node, Ref):
        raise UncalculableError("SUBTOTAL or AGGREGATE of something other than a reference")
    found: list[Scalar] = []
    for sheet in engine.sheets_of(node):
        top = node.top or 1
        last_row, last_column = extent(sheet)
        bottom = node.bottom or max(last_row, top)
        left = node.left or 1
        right = node.right or max(last_column, left)
        engine.charge((bottom - top + 1) * (right - left + 1))
        for row in range(top, bottom + 1):
            is_hidden = sheet.row_dimensions[row].hidden
            if is_hidden and hidden == "skip":
                continue
            if is_hidden and hidden == "unknown":
                raise UncalculableError("hidden rows in SUBTOTAL")
            for col in range(left, right + 1):
                cell = sheet._cells.get((row, col))
                if (
                    skip_nested
                    and cell is not None
                    and cell.data_type == "f"
                    and any(name in str(cell.value).upper() for name in _NESTED)
                ):
                    continue
                value = engine.cell_value(sheet, row, col)
                if skip_errors and isinstance(value, ExcelError):
                    continue
                found.append(value)
    return found


def _apply(name: str, values: list[Scalar], extra: tuple[Value, ...] = ()) -> Value:
    for value in values:
        if isinstance(value, ExcelError):
            raise FormulaError(value)
    if name == "COUNTA":
        return float(sum(v is not None for v in values))
    numbers_only = Grid([[v for v in values if is_number(v)]])
    return FUNCTIONS[name].call(numbers_only, *extra)


@function("SUBTOTAL", kind="lazy")
def subtotal(engine: "Engine", function_num: Node, *references: Node) -> Value:
    code = to_int(engine.scalar(function_num))
    ignore_hidden = code > 100
    if not 1 <= code % 100 <= 11 or code > 111:
        raise FormulaError(VALUE)
    values: list[Scalar] = []
    for reference in references:
        values += _cells(
            engine,
            reference,
            hidden="skip" if ignore_hidden else "unknown",
            skip_nested=True,
            skip_errors=False,
        )
    return _apply(_SUBTOTAL[code % 100], values)


@function("AGGREGATE", kind="lazy")
def aggregate(engine: "Engine", function_num: Node, options: Node, *rest: Node) -> Value:
    code, option = to_int(engine.scalar(function_num)), to_int(engine.scalar(options))
    if not 0 <= option <= 7 or not 1 <= code <= 19:
        raise FormulaError(VALUE)
    skip_nested = option in (0, 1, 2, 3)
    skip_errors = option in (2, 3, 6, 7)
    hidden = "skip" if option in (1, 3, 5, 7) else "keep"
    if code in _WITH_K:
        if len(rest) != 2:
            raise FormulaError(VALUE)
        values = _cells(engine, rest[0], hidden, skip_nested, skip_errors)
        return _apply(_WITH_K[code], values, (engine.scalar(rest[1]),))
    if not rest:
        raise FormulaError(VALUE)
    values = [
        value
        for reference in rest
        for value in _cells(engine, reference, hidden, skip_nested, skip_errors)
    ]
    return _apply(_AGGREGATE[code], values)
