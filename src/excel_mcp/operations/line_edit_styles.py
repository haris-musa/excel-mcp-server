"""The formats of the cells that inserting rows or columns creates."""

# pyright: reportAttributeAccessIssue=false
# openpyxl's stubs lack the style tables of a workbook.

from copy import copy
from typing import cast

from openpyxl.styles import Border, Side
from openpyxl.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from excel_mcp.lazy_workbook import LazyCells, held_cells
from excel_mcp.package.lines import LineEdit


class InheritedStyles:
    """Inserted cells are formatted like the cells above or to the left, as in Excel's default.

    Excel keeps a side border only where the cell on the far side of the insertion has the
    same one, and the borders along the insertion belong to the neighbours. A sheet has few
    kinds of neighbours, so each is worked out once.
    """

    def __init__(self, sheet: Worksheet, edit: LineEdit) -> None:
        self.book = cast(Workbook, sheet.parent)
        self.rows = edit.axis == "rows"
        self.borders: dict[tuple[int, int | None], int] = {}
        self.styles: dict[tuple[int, int | None], int] = {}

    def border_id(self, above: int, after: int | None) -> int:
        """The border of a cell inserted next to a cell with border ``above``; ``after`` is the
        border of the cell on the other side, if there is one."""
        key = (above, after)
        if key not in self.borders:
            table = self.book._borders
            self.borders[key] = table.add(
                _inherited_border(
                    copy(table[above]), None if after is None else copy(table[after]), self.rows
                )
            )
        return self.borders[key]

    def style_id(self, above: int, after: int | None) -> int:
        """`border_id` for cells that are given as style ids, and the id of the whole style."""
        key = (above, after)
        if key not in self.styles:
            table = self.book._cell_styles
            style = copy(table[above])
            style.borderId = self.border_id(
                style.borderId, None if after is None else table[after].borderId
            )
            self.styles[key] = table.add(style)
        return self.styles[key]


def inherit_cell_styles(sheet: Worksheet, edit: LineEdit) -> None:
    """Format the cells inserted next to the cells that are in memory; a large sheet's other
    rows are done as they are renumbered."""
    if edit.delete:
        return
    rows = edit.axis == "rows"
    cells = sheet._cells
    scan = (
        cells.in_row(edit.at - 1)
        if rows and isinstance(cells, LazyCells)
        else list(held_cells(sheet).items())
    )
    inherited = InheritedStyles(sheet, edit)
    beyond = edit.at + edit.count
    for (row, column), cell in scan:
        if (row if rows else column) != edit.at - 1:
            continue
        after = cells.get((beyond, column) if rows else (row, beyond))
        border = inherited.border_id(cell._style.borderId, after._style.borderId if after else None)
        for line in range(edit.at, edit.at + edit.count):
            clone = sheet.cell(line, column) if rows else sheet.cell(row, line)
            clone._style = copy(cell._style)
            clone._style.borderId = border


def _inherited_border(above: Border, after: Border | None, rows: bool) -> Border:
    sides = {name: Side() for name in ("left", "right", "top", "bottom")}
    across = ("left", "right") if rows else ("top", "bottom")
    trailing, leading = ("bottom", "top") if rows else ("right", "left")
    if after is not None:
        for name in across:
            if getattr(above, name).style == getattr(after, name).style:
                sides[name] = getattr(above, name)
        if getattr(above, trailing).style == getattr(after, leading).style:
            sides[trailing] = sides[leading] = getattr(above, trailing)
    return Border(
        left=sides["left"], right=sides["right"], top=sides["top"], bottom=sides["bottom"]
    )
