"""Tools for formatting, merging, layout and rules."""

from typing import Annotated, Literal

from pydantic import Field

from excel_mcp.operations import conditional, formatting, rules
from excel_mcp.operations.calculated import formula_values
from excel_mcp.operations.comparison import FormulaResults
from excel_mcp.operations.conditional import ConditionalFormat
from excel_mcp.operations.formatting import CellFormat
from excel_mcp.operations.layout import SheetLayout, apply_layout
from excel_mcp.operations.rules import DataValidationRule
from excel_mcp.refs import CellRange
from excel_mcp.server.params import RangeRef, SheetName, WorkbookPath
from excel_mcp.server.registry import ToolRegistry
from excel_mcp.workspace import Workspace, get_sheet


def register(tools: ToolRegistry, workspace: Workspace) -> None:
    @tools.writer("Format range")
    def format_range(
        path: WorkbookPath, sheet: SheetName, range: RangeRef, style: CellFormat
    ) -> str:
        """Change the font, fill, borders, alignment or number format of a range."""
        with workspace.edit(path) as workbook:
            formatted = formatting.format_range(
                get_sheet(workbook, sheet), range, style, workspace.limits.max_cells
            )
        return f"Formatted {sheet}!{formatted}."

    @tools.destroyer("Merge or unmerge cells")
    def merge_cells(
        path: WorkbookPath,
        sheet: SheetName,
        range: RangeRef,
        action: Annotated[
            Literal["merge", "unmerge"], Field(description="Merge the range or split it again.")
        ] = "merge",
    ) -> str:
        """Merge a range into one cell, or split a merged range again.

        Merging keeps only the top-left value.
        """
        with workspace.edit(path) as workbook:
            target = get_sheet(workbook, sheet)
            limit = workspace.limits.max_cells
            if action == "merge":
                done = formatting.merge_cells(target, range, limit)
            else:
                done = formatting.unmerge_cells(target, range, limit)
        return f"{action.capitalize()}d {sheet}!{done}."

    @tools.writer("Set sheet layout")
    def set_sheet_layout(path: WorkbookPath, sheet: SheetName, layout: SheetLayout) -> str:
        """Set column widths, row heights, hidden or grouped rows and columns, frozen panes,
        auto filter (on a range or a table, with criteria), tab color, sheet visibility, print
        setup and sheet protection.

        Protection discourages edits in Excel but is not security: it does not stop this
        server, and the password is weakly hashed.
        """
        with workspace.edit(path) as workbook:
            target = get_sheet(workbook, sheet)

            def results(area: CellRange) -> FormulaResults:
                return formula_values(workspace, path, target, area)

            apply_layout(target, layout, results, workspace.limits.max_cells)
        return f"Updated the layout of {sheet!r}."

    @tools.writer("Add conditional format")
    def add_conditional_format(
        path: WorkbookPath, sheet: SheetName, range: RangeRef, rule: ConditionalFormat
    ) -> str:
        """Add a conditional format rule to a range: scales, data bars, icon sets, cell value or
        formula rules, top/bottom, average, duplicates, text, dates, blanks and errors.

        Rules are evaluated in priority order (default: added last); formats that conflict
        go to the first rule met.
        """
        with workspace.edit(path) as workbook:
            target = conditional.add_conditional_format(get_sheet(workbook, sheet), range, rule)
        return f"Added a {rule.type} rule to {sheet}!{target}."

    @tools.writer("Add data validation")
    def add_data_validation(
        path: WorkbookPath, sheet: SheetName, range: RangeRef, rule: DataValidationRule
    ) -> str:
        """Restrict what can be entered in a range: a dropdown list (typed in, or from cells or a
        name), whole numbers, decimals, dates, times, text length or a custom formula, with an
        optional input message and error alert.
        """
        with workspace.edit(path) as workbook:
            target = rules.add_data_validation(get_sheet(workbook, sheet), range, rule)
        return f"Added {rule.type} validation to {sheet}!{target}."
