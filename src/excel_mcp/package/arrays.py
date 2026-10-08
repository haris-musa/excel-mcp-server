"""Array formulas between loading and saving: where their ranges are and what they spilled."""

from typing import cast

from openpyxl import Workbook
from openpyxl.cell.cell import Cell
from openpyxl.worksheet.formula import ArrayFormula
from openpyxl.worksheet.worksheet import Worksheet

from excel_mcp.package.model import CellMark, state_of
from excel_mcp.refs import CellRange, parse_range


def prepare(workbook: Workbook) -> None:
    """Make array formulas consistent before the workbook is saved.

    Inserting or deleting rows moves an array formula's cell but not the range it
    declares, so the range follows the cell. A formula that was removed or replaced takes
    its spilled values with it, as it does in Excel.
    """
    for sheet in workbook.worksheets:
        if isinstance(sheet, Worksheet):
            _realign(sheet)
    for package in state_of(workbook).sheets.values():
        for mark in package.marks:
            if mark.spill and not (_holds(mark.cell) and mark.cell.value is mark.value[0]):
                release(mark)


def release(mark: CellMark) -> None:
    """Clear the cells an array formula spilled into, unless they were changed since."""
    for cell, value in mark.spill:
        if _holds(cell) and (cell.value, cell.data_type) == value:
            cell.value = None
    mark.spill = []


def _holds(cell: Cell) -> bool:
    """Whether the cell is still in its sheet, at the position it has."""
    sheet = cell.parent
    return isinstance(sheet, Worksheet) and sheet._cells.get((cell.row, cell.column)) is cell


def _realign(sheet: Worksheet) -> None:
    for (row, col), cell in sheet._cells.items():
        formula = cell.value
        if isinstance(formula, ArrayFormula):
            area = parse_range(formula.ref)
            if (area.min_row, area.min_col) != (row, col):
                formula.ref = str(_moved(area, row - area.min_row, col - area.min_col))


def _moved(area: CellRange, rows: int, columns: int) -> CellRange:
    return CellRange(
        area.min_row + rows, area.min_col + columns, area.max_row + rows, area.max_col + columns
    )


def copy_marks(source: Worksheet, target: Worksheet) -> None:
    """Give the array formulas of a copied sheet the metadata that the originals have."""
    marks = state_of(cast(Workbook, source.parent)).sheets.get(source)
    package = state_of(cast(Workbook, target.parent)).sheet(target)
    for mark in marks.marks if marks else []:
        copy = target._cells.get((mark.cell.row, mark.cell.column))
        if not (mark.cm or mark.dynamic) or not isinstance(copy, Cell):
            continue
        if not isinstance(copy.value, ArrayFormula):
            continue
        spilled = [
            (cell, (cell.value, cell.data_type))
            for original, _ in mark.spill
            if isinstance(cell := target._cells.get((original.row, original.column)), Cell)
        ]
        package.mark(
            CellMark(
                copy,
                cm=mark.cm,
                dynamic=mark.dynamic,
                value=(copy.value, copy.data_type),
                cached=mark.cached,
                spill=spilled,
            )
        )
