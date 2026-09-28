"""Tools for worksheets and their rows and columns."""

from typing import Annotated

from pydantic import Field

from excel_mcp.operations import inspect, sheets
from excel_mcp.operations.inspect import SheetDetails
from excel_mcp.operations.sheets import Axis
from excel_mcp.server.params import LineCount, LineIndex, SheetName, WorkbookPath
from excel_mcp.server.registry import ToolRegistry
from excel_mcp.workspace import Workspace, get_sheet

NewSheetName = Annotated[
    str, Field(description="New sheet name: 1-31 characters, none of [ ] : * ? / \\.")
]
AxisParam = Annotated[Axis, Field(description="Whether to act on rows or columns.")]


def register(tools: ToolRegistry, workspace: Workspace) -> None:
    @tools.reader("Describe sheet")
    def describe_sheet(path: WorkbookPath, sheet: SheetName) -> SheetDetails:
        """Describe a sheet's structure: used range, frozen panes, merged ranges, tables,
        charts, data validation, conditional formats and custom column widths."""
        with workspace.read(path) as workbook:
            return inspect.describe_sheet(get_sheet(workbook, sheet))

    @tools.writer("Create sheet")
    def create_sheet(
        path: WorkbookPath,
        sheet: NewSheetName,
        position: Annotated[
            int | None, Field(ge=1, description="1-based position. Default: after the last.")
        ] = None,
    ) -> str:
        """Add an empty worksheet."""
        with workspace.edit(path) as workbook:
            sheets.create_sheet(workbook, sheet, position)
        return f"Created sheet {sheet!r}."

    @tools.writer("Rename sheet")
    def rename_sheet(path: WorkbookPath, sheet: SheetName, new_name: NewSheetName) -> str:
        """Rename a worksheet. Formulas that refer to the old name are not updated."""
        with workspace.edit(path) as workbook:
            sheets.rename_sheet(workbook, sheet, new_name)
        return f"Renamed sheet {sheet!r} to {new_name!r}."

    @tools.writer("Copy sheet")
    def copy_sheet(path: WorkbookPath, sheet: SheetName, new_name: NewSheetName) -> str:
        """Duplicate a worksheet (values, styles and dimensions) within the workbook."""
        with workspace.edit(path) as workbook:
            sheets.copy_sheet(workbook, sheet, new_name)
        return f"Copied sheet {sheet!r} to {new_name!r}."

    @tools.destroyer("Delete sheet")
    def delete_sheet(path: WorkbookPath, sheet: SheetName) -> str:
        """Delete a worksheet and everything on it."""
        with workspace.edit(path) as workbook:
            sheets.delete_sheet(workbook, sheet)
        return f"Deleted sheet {sheet!r}."

    @tools.writer("Insert rows or columns")
    def insert_rows_or_columns(
        path: WorkbookPath, sheet: SheetName, axis: AxisParam, at: LineIndex, count: LineCount = 1
    ) -> str:
        """Insert empty rows or columns before position `at`, shifting the rest down or right.

        Formulas, merged ranges, charts and tables that refer to shifted cells are not
        updated, so check them afterwards.
        """
        with workspace.edit(path) as workbook:
            sheets.insert_lines(get_sheet(workbook, sheet), axis, at, count)
        return f"Inserted {count} {axis} at {at} in {sheet!r}."

    @tools.destroyer("Delete rows or columns")
    def delete_rows_or_columns(
        path: WorkbookPath, sheet: SheetName, axis: AxisParam, at: LineIndex, count: LineCount = 1
    ) -> str:
        """Delete rows or columns starting at position `at`, shifting the rest up or left.

        Formulas, merged ranges, charts and tables that refer to shifted cells are not
        updated, so check them afterwards.
        """
        with workspace.edit(path) as workbook:
            sheets.delete_lines(get_sheet(workbook, sheet), axis, at, count)
        return f"Deleted {count} {axis} starting at {at} in {sheet!r}."
