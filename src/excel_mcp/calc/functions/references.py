"""Functions about references: ISREF, ISFORMULA, FORMULATEXT, SHEET, SHEETS, AREAS, ADDRESS."""

import re
from typing import TYPE_CHECKING

from openpyxl.utils import get_column_letter
from openpyxl.worksheet.formula import ArrayFormula

from excel_mcp.calc.parser import Call, Node, Ref
from excel_mcp.calc.registry import function
from excel_mcp.calc.values import (
    NA,
    VALUE,
    ExcelError,
    FormulaError,
    Grid,
    RefGrid,
    Scalar,
    UncalculableError,
    Value,
    to_bool,
    to_int,
    to_text,
)
from excel_mcp.refs import MAX_COLUMN, MAX_ROW

if TYPE_CHECKING:
    from excel_mcp.calc.engine import Engine

_PLAIN_SHEET_NAME = re.compile(r"[^\W\d][\w.]*", re.UNICODE)
_CELL_LIKE = re.compile(r"(?:[A-Za-z]{1,3}\d+|[Rr](?:\d+)?[Cc](?:\d+)?|[Rr]|[Cc])", re.UNICODE)


def _reference(engine: "Engine", node: Node) -> RefGrid:
    value = engine.eval(node)
    if isinstance(value, ExcelError):
        raise FormulaError(value)
    if not isinstance(value, RefGrid):
        raise FormulaError(VALUE)
    return value


@function("ISREF", kind="lazy")
def isref(engine: "Engine", node: Node) -> bool:
    value = engine.eval(node)
    if isinstance(node, Call) and node.name == "ANCHORARRAY":
        return isinstance(value, Grid)
    return isinstance(value, RefGrid)


@function("ISFORMULA", kind="lazy")
def isformula(engine: "Engine", node: Node) -> Value:
    found = _reference(engine, node)
    if found.clipped_rows or found.clipped_cols:
        raise UncalculableError("ISFORMULA of a whole row or column")
    sheet = engine.find_sheet(found.sheet or engine.here.sheet.title)
    rows = [
        [
            ((cell := sheet._cells.get((found.top + r, found.left + c))) is not None)
            and cell.data_type == "f"
            for c in range(found.width)
        ]
        for r in range(found.height)
    ]
    return rows[0][0] if found.height == found.width == 1 else Grid(rows)


@function("FORMULATEXT", kind="lazy")
def formulatext(engine: "Engine", node: Node) -> str:
    found = _reference(engine, node)
    sheet = engine.find_sheet(found.sheet or engine.here.sheet.title)
    cell = sheet._cells.get((found.top, found.left))
    if cell is None or cell.data_type != "f":
        raise FormulaError(NA)
    formula = str(cell.value.text if isinstance(cell.value, ArrayFormula) else cell.value)
    return re.sub(r"_xl\w+\.", "", formula)


@function("SHEET", kind="lazy")
def sheet_number(engine: "Engine", node: Node | None = None) -> float:
    order = [sheet.title for sheet in engine.sheets.values()]
    if node is None:
        return float(order.index(engine.here.sheet.title) + 1)
    if isinstance(node, Ref):
        return float(order.index(engine.sheets_of(node)[0].title) + 1)
    value = engine.eval(node)
    if isinstance(value, RefGrid):
        return _first_sheet_position(engine, value)
    if isinstance(value, str):
        names = [name.casefold() for name in order]
        if value.casefold() not in names:
            raise FormulaError(NA)
        return float(names.index(value.casefold()) + 1)
    raise FormulaError(NA if isinstance(value, ExcelError) else VALUE)


def _first_sheet_position(engine: "Engine", grid: RefGrid) -> float:
    order = [sheet.title for sheet in engine.sheets.values()]
    if grid.sheet is None:
        raise UncalculableError("SHEET of a 3-D reference")
    return float(order.index(grid.sheet) + 1)


@function("SHEETS", kind="lazy")
def sheets(engine: "Engine", node: Node | None = None) -> float:
    if node is None:
        return float(len(engine.sheets))
    return float(_reference(engine, node).sheets)


@function("AREAS", kind="lazy")
def areas(engine: "Engine", node: Node) -> float:
    if isinstance(node, Call) and node.name == "UNION":
        return sum(areas(engine, member) for member in node.args)
    _reference(engine, node)
    return 1.0


def _quote_sheet(name: str) -> str:
    if not name or (_PLAIN_SHEET_NAME.fullmatch(name) and not _CELL_LIKE.fullmatch(name)):
        return name
    return "'" + name.replace("'", "''") + "'"


@function("ADDRESS", kind="scalar")
def address(
    row: Scalar,
    column: Scalar,
    absolute: Scalar = 1,
    a1: Scalar = True,
    sheet: Scalar = ...,  # pyright: ignore[reportArgumentType]
) -> str:
    r, c = to_int(row), to_int(column)
    style = 1 if absolute is None else to_int(absolute)
    if not (1 <= r <= MAX_ROW and 1 <= c <= MAX_COLUMN and 1 <= style <= 4):
        raise FormulaError(VALUE)
    row_fixed, column_fixed = style in (1, 2), style in (1, 3)
    if a1 is None or to_bool(a1):
        text = ("$" if column_fixed else "") + get_column_letter(c)
        text += ("$" if row_fixed else "") + str(r)
    else:
        text = f"R{r}" if row_fixed else f"R[{r}]"
        text += f"C{c}" if column_fixed else f"C[{c}]"
    return text if sheet is ... else f"{_quote_sheet(to_text(sheet))}!{text}"
