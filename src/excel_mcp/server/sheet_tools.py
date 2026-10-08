"""Tools for worksheets and their rows and columns."""

from typing import Annotated

from pydantic import Field

from excel_mcp.operations import inspect, sheets
from excel_mcp.operations.inspect import SheetDetails
from excel_mcp.operations.sheets import Axis
from excel_mcp.server.params import LineCount, LineIndex, SheetName, WorkbookPath
from excel_mcp.server.registry import ToolRegistry
from excel_mcp.workspace import Workspace, get_sheet

AxisParam = Annotated[Axis, Field(description="Rows or columns.")]
NewSheetName = Annotated[
    str, Field(description="New sheet name: 1-31 characters, none of [ ] : * ? / \\.")
]


def register(tools: ToolRegistry, workspace: Workspace) -> None:
    @tools.reader("Describe sheet")
    def describe_sheet(path: WorkbookPath, sheet: SheetName) -> SheetDetails:
        """Describe a sheet's used range, frozen panes, merged ranges, tables, charts, PivotTables,
        images, notes, validation, conditional formats, custom column widths, hidden rows and
        columns, print area and protection. Empty items are omitted.

        Loads the whole workbook into memory, so it is slow on very large files."""
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
        """Duplicate a worksheet (values, styles, dimensions)."""
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
        """Insert empty rows or columns before position `at`.

        References in formulas, merged ranges, charts and tables are not updated.
        """
        with workspace.edit(path) as workbook:
            sheets.insert_lines(get_sheet(workbook, sheet), axis, at, count)
        return f"Inserted {count} {axis} at {at} in {sheet!r}."

    @tools.destroyer("Delete rows or columns")
    def delete_rows_or_columns(
        path: WorkbookPath, sheet: SheetName, axis: AxisParam, at: LineIndex, count: LineCount = 1
    ) -> str:
        """Delete rows or columns starting at position `at`.

        References in formulas, merged ranges, charts and tables are not updated.
        """
        with workspace.edit(path) as workbook:
            sheets.delete_lines(get_sheet(workbook, sheet), axis, at, count)
        return f"Deleted {count} {axis} starting at {at} in {sheet!r}."
