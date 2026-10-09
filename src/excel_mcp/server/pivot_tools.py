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
        sheet: Annotated[SheetName, Field(description="Sheet for the PivotTable.")],
        at: Annotated[CellRef, Field(description="Top-left cell; the area must be empty.")],
        source: Annotated[
            RangeRef,
            Field(
                description="Header row of unique text labels, then records, e.g. "
                "'Data!A1:E200' (no sheet: `sheet`). Each column holds only text, numbers or "
                "dates (blanks fine), not formulas."
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
            Field(max_length=64, description="Headers across the top, outermost first."),
        ] = [],  # noqa: B006
        filter_fields: Annotated[
            list[str],
            Field(max_length=64, description="Headers as page filters."),
        ] = [],  # noqa: B006
        field_settings: Annotated[
            list[PivotField],
            Field(
                max_length=64,
                description="Items to show, sort order and grouping for row, column or "
                "filter headers.",
            ),
        ] = [],  # noqa: B006
        calculated_fields: Annotated[
            list[CalculatedField],
            Field(max_length=64, description="Usable in value_fields."),
        ] = [],  # noqa: B006
        layout: Annotated[Layout, Field(description="Row label layout.")] = "tabular",
        subtotals: Annotated[bool, Field(description="Of outer fields.")] = True,
        values_in: Annotated[
            ValuesIn, Field(description="Where several value fields go.")
        ] = "columns",
        name: Annotated[
            str | None, Field(description="PivotTable name. Default: PivotTableN.")
        ] = None,
    ) -> Changed:
        """Add a PivotTable that summarizes a block of data. Returns its name and cells.

        Excel can refresh it; its figures (with subtotals and Grand Totals) are also
        written into the cells. A field can be used only once among row, column and filter
        fields.
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
        name: Annotated[str, Field(description="Name from describe_sheet.")],
    ) -> Changed:
        """Remove a PivotTable and clear its cells; the source stays. Fails while slicers or
        timelines are connected to it."""
        with workspace.edit(path) as workbook:
            removed = pivot_index.delete_pivot(
                get_sheet(workbook, sheet), name, workspace.limits.max_cells
            )
        return Changed(sheet=sheet, name=removed.name, range=removed.range)
