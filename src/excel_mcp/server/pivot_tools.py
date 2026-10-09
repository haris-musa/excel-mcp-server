"""Tools for PivotTables."""

from typing import Annotated

from pydantic import Field

from excel_mcp.operations import pivot, pivot_index
from excel_mcp.operations.charts_data import split_sheet
from excel_mcp.operations.pivot import PivotRequest
from excel_mcp.operations.pivot_options import (
    CalculatedField,
    Layout,
    PivotField,
    PivotValue,
    ValuesIn,
)
from excel_mcp.server.params import CellRef, RangeRef, SheetName, WorkbookPath
from excel_mcp.server.registry import ToolRegistry
from excel_mcp.server.results import Changed
from excel_mcp.workspace import Workspace, get_sheet


def register(tools: ToolRegistry, workspace: Workspace) -> None:
    @tools.writer("Create pivot table")
    def create_pivot_table(
        path: WorkbookPath,
        sheet: Annotated[SheetName, Field(description="Sheet to put the PivotTable on.")],
        at: Annotated[
            CellRef, Field(description="Top-left cell of the PivotTable; the area must be empty.")
        ],
        source: Annotated[
            RangeRef,
            Field(
                description="Header row of unique text labels, then one record per row, e.g. "
                "'Data!A1:E200' (`sheet` if no sheet is given). Each column holds only text, "
                "only numbers or only dates (blanks are fine), not formulas."
            ),
        ],
        row_fields: Annotated[
            list[str],
            Field(
                min_length=1,
                max_length=64,
                description="Headers to group by down the side, outermost first.",
            ),
        ],
        value_fields: Annotated[
            list[PivotValue],
            Field(min_length=1, max_length=64, description="Headers to summarize."),
        ],
        column_fields: Annotated[
            list[str],
            Field(max_length=64, description="Headers to spread across the top, outermost first."),
        ] = [],  # noqa: B006
        filter_fields: Annotated[
            list[str],
            Field(max_length=64, description="Headers offered as page filters above the table."),
        ] = [],  # noqa: B006
        field_settings: Annotated[
            list[PivotField],
            Field(
                max_length=64,
                description="Per-field settings for headers used in row, column or filter fields: "
                "items to show (`show_items`), sort order, date or number grouping.",
            ),
        ] = [],  # noqa: B006
        calculated_fields: Annotated[
            list[CalculatedField],
            Field(max_length=64, description="Fields calculated from others, usable in values."),
        ] = [],  # noqa: B006
        layout: Annotated[
            Layout, Field(description="Report layout of the row labels.")
        ] = "tabular",
        subtotals: Annotated[bool, Field(description="Show subtotals of outer fields.")] = True,
        values_in: Annotated[
            ValuesIn, Field(description="Where several value fields go: as columns or as rows.")
        ] = "columns",
        name: Annotated[
            str | None, Field(description="PivotTable name. Default: PivotTableN.")
        ] = None,
    ) -> Changed:
        """Add an Excel PivotTable that summarizes a block of data. Returns its name and cells.

        Excel can refresh it (Data > Refresh All) when the source changes, and shows the same
        figures. They are also written into the cells, laid out as Excel does (subtotals, Grand
        Total row and column), so other tools can read them; leave the Grand Total out of a
        chart's source. A field can be used only once among row, column and filter fields;
        filters show every item unless `field_settings` sets `show_items`. Items that tie when
        sorted by value may swap places when Excel refreshes. describe_sheet lists PivotTables;
        delete_pivot_table removes one.
        """
        request = PivotRequest(
            row_fields,
            column_fields,
            value_fields,
            filter_fields,
            field_settings,
            calculated_fields,
            layout,
            subtotals,
            values_in,
            name,
        )
        with workspace.edit(path) as workbook:
            source_title, cells = split_sheet(workbook, sheet, source)
            pivot_name, area = pivot.create_pivot(
                workbook,
                get_sheet(workbook, source_title),
                cells,
                get_sheet(workbook, sheet),
                at,
                request,
                workspace.limits.max_cells,
            )
        return Changed(sheet=sheet, name=pivot_name, range=str(area))

    @tools.destroyer("Delete pivot table")
    def delete_pivot_table(
        path: WorkbookPath,
        sheet: SheetName,
        name: Annotated[str, Field(description="PivotTable name, as listed by describe_sheet.")],
    ) -> Changed:
        """Remove a PivotTable and clear the cells it fills. The source data is left untouched.

        Fails while slicers or timelines are connected to it.
        """
        with workspace.edit(path) as workbook:
            removed = pivot_index.delete_pivot(
                get_sheet(workbook, sheet), name, workspace.limits.max_cells
            )
        return Changed(sheet=sheet, name=removed.name, range=removed.range)
