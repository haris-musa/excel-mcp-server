"""Tools for PivotTables."""

from typing import Annotated

from pydantic import Field

from excel_mcp.operations import pivot, pivot_index
from excel_mcp.operations.pivot import PivotValue
from excel_mcp.server.params import CellRef, RangeRef, SheetName, WorkbookPath
from excel_mcp.server.registry import ToolRegistry
from excel_mcp.workspace import Workspace, get_sheet


def register(tools: ToolRegistry, workspace: Workspace) -> None:
    @tools.writer("Create pivot table")
    def create_pivot_table(
        path: WorkbookPath,
        source_sheet: SheetName,
        source_range: Annotated[
            RangeRef,
            Field(
                description="Data to summarize: a header row of unique text labels, then one "
                "record per row, e.g. 'A1:E200'. Columns must hold only text, only numbers "
                "or only dates (blanks are fine), not formulas."
            ),
        ],
        rows: Annotated[
            list[str],
            Field(min_length=1, description="Headers to group by down the side, outermost first."),
        ],
        values: Annotated[
            list[PivotValue], Field(min_length=1, description="Headers to summarize, with how.")
        ],
        target_sheet: SheetName,
        target_cell: Annotated[
            CellRef, Field(description="Top-left cell of the PivotTable; the area must be empty.")
        ],
        columns: Annotated[
            list[str], Field(description="Headers to spread across the top, outermost first.")
        ] = [],  # noqa: B006
        filters: Annotated[
            list[str], Field(description="Headers to offer as page filters above the table.")
        ] = [],  # noqa: B006
        name: Annotated[
            str | None, Field(description="PivotTable name. Default: PivotTableN.")
        ] = None,
    ) -> str:
        """Add a real Excel PivotTable that summarizes a block of data.

        Excel shows it at once and can refresh it (Data > Refresh All) after the source
        changes, and you can drag fields or change summaries there. Values are also
        written into the cells so other tools can read them: with subtotals for outer
        row fields and grand totals. Filters start out showing everything. A field can
        be used only once among rows, columns and filters. Remove a PivotTable with
        delete_pivot_table; describe_sheet lists them.
        """
        with workspace.edit(path) as workbook:
            pivot_name, area = pivot.create_pivot(
                workbook,
                get_sheet(workbook, source_sheet),
                source_range,
                get_sheet(workbook, target_sheet),
                target_cell,
                rows,
                columns,
                values,
                filters,
                name,
                workspace.limits.max_cells,
            )
        return f"Created PivotTable {pivot_name!r} at {target_sheet}!{area}."

    @tools.destroyer("Delete pivot table")
    def delete_pivot_table(
        path: WorkbookPath,
        sheet: SheetName,
        name: Annotated[str, Field(description="PivotTable name, as listed by describe_sheet.")],
    ) -> str:
        """Remove a PivotTable and clear the cells it fills. The source data is left untouched."""
        with workspace.edit(path) as workbook:
            removed = pivot_index.delete_pivot(
                get_sheet(workbook, sheet), name, workspace.limits.max_cells
            )
        return f"Deleted PivotTable {removed.name!r} ({removed.range}) from {sheet}."
