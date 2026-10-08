"""Tools for worksheets and their rows and columns."""

from typing import Annotated

from pydantic import Field

from excel_mcp.operations import inspect, line_edits, sheet_copy, sheets
from excel_mcp.operations.calculated import formula_values
from excel_mcp.operations.inspect import SheetDetails
from excel_mcp.package.lines import Axis
from excel_mcp.server.params import LineCount, LineIndex, SheetName, WorkbookPath
from excel_mcp.server.registry import ToolRegistry
from excel_mcp.server.results import Changed
from excel_mcp.workspace import Workspace, get_sheet

AxisParam = Axis
NewSheetName = Annotated[str, Field(description="1-31 characters, none of [ ] : * ? / \\.")]


def register(tools: ToolRegistry, workspace: Workspace) -> None:
    @tools.reader("Describe sheet")
    def describe_sheet(path: WorkbookPath, sheet: SheetName) -> SheetDetails:
        """Describe a sheet: used range, panes, merges and everything on it (tables, charts,
        PivotTables, slicers, images, notes, hyperlinks, validation, conditional formats,
        sparklines...) with names and cells. Slow on very large files."""
        with workspace.read(path, with_package=True) as workbook:
            return inspect.describe_sheet(get_sheet(workbook, sheet))

    @tools.writer("Create sheet")
    def create_sheet(
        path: WorkbookPath,
        new_name: NewSheetName,
        position: Annotated[int | None, Field(ge=1, description="1-based. Default: last.")] = None,
    ) -> Changed:
        """Add an empty worksheet."""
        with workspace.edit(path) as workbook:
            sheets.create_sheet(workbook, new_name, position)
        return Changed(sheet=new_name)

    @tools.writer("Rename sheet")
    def rename_sheet(path: WorkbookPath, sheet: SheetName, new_name: NewSheetName) -> Changed:
        """Rename a worksheet; references to it are updated as in Excel."""
        with workspace.edit(path) as workbook:
            sheets.rename_sheet(workbook, sheet, new_name)
        return Changed(sheet=new_name)

    @tools.writer("Copy sheet")
    def copy_sheet(path: WorkbookPath, sheet: SheetName, new_name: NewSheetName) -> Changed:
        """Copy a worksheet to a new last sheet, with everything on it, as Excel does.
        Self-references point at the copy; tables get new names (Sales2); PivotTables share
        the original's data."""
        with workspace.edit(path) as workbook:
            skipped = sheet_copy.copy_sheet(
                workbook, sheet, new_name, workspace.limits.max_copy_cells
            )
        note = f"Not copied: {', '.join(skipped)}." if skipped else None
        return Changed(sheet=new_name, note=note)

    @tools.destroyer("Delete sheet")
    def delete_sheet(path: WorkbookPath, sheet: SheetName) -> Changed:
        """Delete a sheet and everything on it. Fails while slicers elsewhere use its
        PivotTables or tables."""
        with workspace.edit(path) as workbook:
            sheets.delete_sheet(workbook, sheet)
        return Changed(sheet=sheet)

    @tools.writer("Insert rows or columns")
    def insert_rows_or_columns(
        path: WorkbookPath,
        sheet: SheetName,
        axis: AxisParam,
        start: LineIndex,
        count: LineCount = 1,
    ) -> Changed:
        """Insert empty rows or columns before `start`. All references move as in Excel;
        edits Excel refuses (through an array formula, PivotTable or table header) fail."""
        with workspace.edit(path) as workbook:
            target = get_sheet(workbook, sheet)
            line_edits.edit_lines(
                workbook,
                target,
                axis,
                start,
                count,
                delete=False,
                formula_values=lambda area: formula_values(workspace, path, target, area),
            )
        return Changed(sheet=sheet, range=sheets.line_span(axis, start, count))

    @tools.destroyer("Delete rows or columns")
    def delete_rows_or_columns(
        path: WorkbookPath,
        sheet: SheetName,
        axis: AxisParam,
        start: LineIndex,
        count: LineCount = 1,
    ) -> Changed:
        """Delete rows or columns from `start`. References move as in insert_rows_or_columns;
        one to a deleted cell becomes #REF!."""
        with workspace.edit(path) as workbook:
            target = get_sheet(workbook, sheet)
            line_edits.edit_lines(
                workbook,
                target,
                axis,
                start,
                count,
                delete=True,
                formula_values=lambda area: formula_values(workspace, path, target, area),
            )
        return Changed(sheet=sheet, range=sheets.line_span(axis, start, count))
