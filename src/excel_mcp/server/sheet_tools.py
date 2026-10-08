"""Tools for worksheets and their rows and columns."""

from typing import Annotated

from pydantic import Field

from excel_mcp.operations import inspect, sheet_copy, sheets
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
        images, notes, threaded comments, hyperlinks, validation, conditional formats, custom
        column widths, hidden rows and columns, print area and protection. Empty items are
        omitted.

        Loads the whole workbook into memory, so it is slow on very large files."""
        with workspace.read(path, with_package=True) as workbook:
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
        """Copy a worksheet to a new sheet at the end, as Excel's "Create a copy" does.

        Copies cells, styles, merges, sizes, hidden rows and columns, freeze panes, filters, data
        validation, conditional formats, images, notes, charts, tables, PivotTables, print setup,
        protection and sheet-scoped names. References to the sheet itself, including chart data,
        point at the copy. Tables get new names (Sales becomes Sales2). PivotTables share the
        original's data. Workbook-scoped names are not duplicated.
        """
        with workspace.edit(path) as workbook:
            sheet_copy.copy_sheet(workbook, sheet, new_name)
        return f"Copied sheet {sheet!r} to {new_name!r}."

    @tools.destroyer("Delete sheet")
    def delete_sheet(path: WorkbookPath, sheet: SheetName) -> str:
        """Delete a worksheet or chart sheet and everything on it.

        Fails while slicers on other sheets use its PivotTables or tables.
        """
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
        return f"Inserted {sheets.describe_lines(axis, at, count)}."

    @tools.destroyer("Delete rows or columns")
    def delete_rows_or_columns(
        path: WorkbookPath, sheet: SheetName, axis: AxisParam, at: LineIndex, count: LineCount = 1
    ) -> str:
        """Delete rows or columns starting at position `at`.

        References in formulas, merged ranges, charts and tables are not updated.
        """
        with workspace.edit(path) as workbook:
            sheets.delete_lines(get_sheet(workbook, sheet), axis, at, count)
        return f"Deleted {sheets.describe_lines(axis, at, count)}."
