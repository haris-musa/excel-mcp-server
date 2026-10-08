"""Tools for worksheets and their rows and columns."""

from typing import Annotated

from pydantic import Field

from excel_mcp.operations import inspect, line_edits, sheet_copy, sheets
from excel_mcp.operations.calculated import formula_values
from excel_mcp.operations.inspect import SheetDetails
from excel_mcp.package.lines import Axis
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
        slicers, timelines, images, notes, hyperlinks, validation, conditional formats, sparklines,
        custom column widths, hidden rows and columns, print area and protection. Empty items are
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
        """Rename a worksheet. References to it are updated as in Excel: formulas, names, rules,
        charts and PivotTable sources."""
        with workspace.edit(path) as workbook:
            sheets.rename_sheet(workbook, sheet, new_name)
        return f"Renamed sheet {sheet!r} to {new_name!r}."

    @tools.writer("Copy sheet")
    def copy_sheet(path: WorkbookPath, sheet: SheetName, new_name: NewSheetName) -> str:
        """Copy a worksheet to a new sheet at the end, as Excel's "Create a copy" does.

        Copies cells, styles, merges, sizes, hidden rows and columns, freeze panes, filters, data
        validation, conditional formats, images, notes, charts, tables, PivotTables, print setup,
        protection, slicers, timelines and sheet-scoped names. References to the sheet itself,
        including chart data, point at the copy. Tables get new names (Sales becomes Sales2).
        PivotTables share the original's data. Workbook-scoped names are not duplicated.
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

        Like Excel, every reference moves: formulas on all sheets, names, conditional formats,
        validation, merged cells, tables, charts, PivotTables, filters and print settings.
        An edit Excel refuses (through an array formula, a PivotTable or a table header)
        fails.
        """
        with workspace.edit(path) as workbook:
            target = get_sheet(workbook, sheet)
            line_edits.edit_lines(
                workbook,
                target,
                axis,
                at,
                count,
                delete=False,
                formula_values=lambda area: formula_values(workspace, path, target, area),
            )
        return f"Inserted {sheets.describe_lines(axis, at, count)}."

    @tools.destroyer("Delete rows or columns")
    def delete_rows_or_columns(
        path: WorkbookPath, sheet: SheetName, axis: AxisParam, at: LineIndex, count: LineCount = 1
    ) -> str:
        """Delete rows or columns starting at position `at`.

        Like Excel, every reference moves, and one to a deleted cell becomes #REF!. See
        insert_rows_or_columns.
        """
        with workspace.edit(path) as workbook:
            target = get_sheet(workbook, sheet)
            line_edits.edit_lines(
                workbook,
                target,
                axis,
                at,
                count,
                delete=True,
                formula_values=lambda area: formula_values(workspace, path, target, area),
            )
        return f"Deleted {sheets.describe_lines(axis, at, count)}."
