"""How a sheet looks when opened: zoom, gridlines, headings, selection, the active sheet."""

from typing import Literal, cast

from openpyxl.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet
from pydantic import BaseModel, Field

from excel_mcp.errors import InvalidArgumentError
from excel_mcp.inputs import InputModel
from excel_mcp.refs import parse_cell


class ViewOptions(InputModel):
    zoom: int | None = Field(default=None, ge=10, le=400)
    gridlines: bool | None = None
    headings: bool | None = Field(default=None, description="Row and column labels.")
    show_formulas: bool | None = None
    right_to_left: bool | None = None
    active: Literal[True] | None = Field(
        default=None, description="Show this sheet when the file opens."
    )
    selected_cell: str | None = Field(default=None, description="e.g. 'B2'.")


class ViewInfo(BaseModel):
    zoom: int = 100
    gridlines: bool = True
    headings: bool = True
    show_formulas: bool = False
    right_to_left: bool = False
    active: bool = False
    selected_cell: str = "A1"


def apply_view(sheet: Worksheet, options: ViewOptions) -> None:
    view = sheet.sheet_view
    if options.zoom is not None:
        view.zoomScale = view.zoomScaleNormal = options.zoom
    if options.gridlines is not None:
        view.showGridLines = options.gridlines
    if options.headings is not None:
        view.showRowColHeaders = options.headings
    if options.show_formulas is not None:
        view.showFormulas = options.show_formulas
    if options.right_to_left is not None:
        view.rightToLeft = options.right_to_left
    if options.selected_cell is not None:
        parse_cell(options.selected_cell)
        # With frozen panes the last selection belongs to the pane that has the cursor.
        selection = view.selection[-1]
        selection.activeCell = selection.sqref = options.selected_cell.upper()
    if options.active:
        _activate(sheet)


def read_view(sheet: Worksheet) -> ViewInfo:
    view = sheet.sheet_view
    return ViewInfo(
        zoom=view.zoomScale or 100,
        gridlines=view.showGridLines is not False,
        headings=view.showRowColHeaders is not False,
        show_formulas=bool(view.showFormulas),
        right_to_left=bool(view.rightToLeft),
        active=cast(Workbook, sheet.parent).active is sheet,
        selected_cell=view.selection[-1].activeCell or "A1",
    )


def move_sheet(sheet: Worksheet, position: int) -> None:
    workbook = cast(Workbook, sheet.parent)
    names = workbook.sheetnames
    if position > len(names):
        raise InvalidArgumentError(f"position must be 1 to {len(names)}, the number of sheets.")
    active = cast(Worksheet, workbook.active).title
    workbook.move_sheet(sheet, position - 1 - names.index(sheet.title))
    workbook.active = workbook.sheetnames.index(active)


def _activate(sheet: Worksheet) -> None:
    if sheet.sheet_state != "visible":
        raise InvalidArgumentError("A hidden sheet cannot be the active sheet; show it first.")
    workbook = cast(Workbook, sheet.parent)
    workbook.active = sheet
    for other in workbook.worksheets:
        other.sheet_view.tabSelected = other is sheet
