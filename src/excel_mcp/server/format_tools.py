"""Tools for formatting, merging, layout and rules."""

from typing import Annotated, Literal

from pydantic import Field

from excel_mcp.operations import conditional, formatting, rules
from excel_mcp.operations.calculated import formula_values
from excel_mcp.operations.comparison import FormulaResults
from excel_mcp.operations.conditional_rule import ConditionalFormat
from excel_mcp.operations.formatting import CellFormat
from excel_mcp.operations.layout import SheetLayout, apply_layout
from excel_mcp.operations.rules import DataValidationRule
from excel_mcp.refs import CellRange
from excel_mcp.server.params import RangeRef, SheetName, WorkbookPath
from excel_mcp.server.registry import ToolRegistry
from excel_mcp.server.results import Changed
from excel_mcp.workspace import Workspace, get_sheet


def register(tools: ToolRegistry, workspace: Workspace) -> None:
    @tools.writer("Format range")
    def format_range(
        path: WorkbookPath, sheet: SheetName, range: RangeRef, style: CellFormat
    ) -> Changed:
        """Change the font, fill, borders, alignment, number format or protection flags of a
        range. `locked` and `formula_hidden` apply once the sheet is protected."""
        with workspace.edit(path) as workbook:
            formatted = formatting.format_range(
                get_sheet(workbook, sheet), range, style, workspace.limits.max_cells
            )
        return Changed(sheet=sheet, range=formatted)

    @tools.destroyer("Merge or unmerge cells")
    def merge_cells(
        path: WorkbookPath,
        sheet: SheetName,
        range: RangeRef,
        action: Annotated[
            Literal["merge", "unmerge"], Field(description="Unmerge splits it again.")
        ] = "merge",
    ) -> Changed:
        """Merge a range into one cell (keeping the top-left value), or unmerge it."""
        with workspace.edit(path) as workbook:
            target = get_sheet(workbook, sheet)
            limit = workspace.limits.max_cells
            if action == "merge":
                done = formatting.merge_cells(target, range, limit)
            else:
                done = formatting.unmerge_cells(target, range, limit)
        return Changed(sheet=sheet, range=done)

    @tools.writer("Set sheet layout")
    def set_sheet_layout(path: WorkbookPath, sheet: SheetName, layout: SheetLayout) -> Changed:
        """Set column widths, row heights, hidden or grouped lines, frozen panes, auto
        filter, tab color, visibility, position, view options, print setup and protection.

        Protection discourages edits in Excel but is not security (it does not stop this
        server; the password is weakly hashed).
        """
        with workspace.edit(path) as workbook:
            target = get_sheet(workbook, sheet)

            def results(area: CellRange) -> FormulaResults:
                return formula_values(workspace, path, target, area)

            apply_layout(target, layout, results, workspace.limits.max_cells)
        return Changed(sheet=sheet)

    @tools.writer("Add conditional format")
    def add_conditional_format(
        path: WorkbookPath, sheet: SheetName, range: RangeRef, rule: ConditionalFormat
    ) -> Changed:
        """Add a conditional format rule to a range. Rules apply in priority order (default:
        last); of conflicting formats the first rule met wins."""
        with workspace.edit(path) as workbook:
            target = conditional.add_conditional_format(get_sheet(workbook, sheet), range, rule)
        return Changed(sheet=sheet, range=target)

    @tools.writer("Add data validation")
    def add_data_validation(
        path: WorkbookPath, sheet: SheetName, range: RangeRef, rule: DataValidationRule
    ) -> Changed:
        """Restrict what can be entered in a range (list dropdown, numbers, dates, times,
        text length or custom formula), with optional input message and error alert."""
        with workspace.edit(path) as workbook:
            target = rules.add_data_validation(get_sheet(workbook, sheet), range, rule)
        return Changed(sheet=sheet, range=target)
