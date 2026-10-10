"""Tools for slicers and timelines."""

from typing import Annotated, Literal

from pydantic import Field

from excel_mcp.operations import slicer_manage, slicers
from excel_mcp.operations.calculated import formula_values
from excel_mcp.operations.slicers import SlicerRequest, Source, Timeline
from excel_mcp.server.params import CellRef, SheetName, WorkbookPath
from excel_mcp.server.registry import ToolRegistry
from excel_mcp.server.results import Changed
from excel_mcp.workspace import Workspace, get_sheet

_STALE = (
    "The PivotTable was not made by create_pivot_table, so its figures are not recalculated "
    "here: Excel recalculates them when the file is opened."
)


def register(tools: ToolRegistry, workspace: Workspace) -> None:
    @tools.writer("Add slicer")
    def add_slicer(
        path: WorkbookPath,
        sheet: Annotated[SheetName, Field(description="Sheet for the slicer.")],
        target: Annotated[Source, Field(description="Table or PivotTable to filter.")],
        field: Annotated[
            str, Field(description="Column header or field, e.g. 'Months (Date)' for a date group.")
        ],
        at: CellRef,
        width_cm: Annotated[
            float | None, Field(gt=0, le=100, description="Default: 5.1 (timeline: 9.3).")
        ] = None,
        height_cm: Annotated[
            float | None, Field(gt=0, le=100, description="Default: 7.4 (timeline: 3.8).")
        ] = None,
        caption: Annotated[str | None, Field(description="Default: the field.")] = None,
        name: Annotated[str | None, Field(description="Default: the field name.")] = None,
        columns: Annotated[int, Field(ge=1, le=20, description="Button columns.")] = 1,
        style: Annotated[
            str | None,
            Field(
                pattern=r"^(Time)?SlicerStyle(Light|Dark|Other)\d$",
                description="e.g. SlicerStyleDark2, TimeSlicerStyleLight1 (timeline).",
            ),
        ] = None,
        selected_items: Annotated[
            list[str], Field(max_length=10_000, description="Items to show. Default: all.")
        ] = [],  # noqa: B006
        sort: Annotated[
            Literal["ascending", "descending"], Field(description="Item order.")
        ] = "ascending",
        hide_empty_items: Annotated[bool, Field(description="Hide items without data.")] = False,
        connect: Annotated[
            list[Source],
            Field(
                max_length=50,
                description="More PivotTables sharing the target's data cache.",
            ),
        ] = [],  # noqa: B006
        timeline: Annotated[
            Timeline | None,
            Field(description="Make a timeline instead, for a date field of a PivotTable."),
        ] = None,
    ) -> Changed:
        """Add a slicer (Insert > Slicer) or timeline that filters a PivotTable or table.

        `selected_items` filters as clicking buttons does. A PivotTable not made by
        create_pivot_table keeps its old figures until Excel recalculates on open.
        """
        request = SlicerRequest(
            target, connect, field, at, width_cm, height_cm, caption, name, columns, style,
            selected_items, sort, hide_empty_items, timeline,
        )  # fmt: skip
        with workspace.edit(path) as workbook:

            def results(owner, area):
                return formula_values(workspace, path, owner, area)

            made, stale = slicers.add_slicer(
                workbook, get_sheet(workbook, sheet), request, results, workspace.limits.max_cells
            )
            area = slicer_manage.placed_range(workbook, get_sheet(workbook, sheet), made)
        return Changed(sheet=sheet, name=made, range=area, note=_STALE if stale else None)

    @tools.destroyer("Delete slicer")
    def delete_slicer(
        path: WorkbookPath,
        sheet: SheetName,
        name: Annotated[str, Field(description="Name from describe_sheet.")],
    ) -> Changed:
        """Remove a slicer or timeline. As in Excel, tables and PivotTable fields stay filtered;
        a timeline's period is cleared."""
        with workspace.edit(path) as workbook:
            slicer_manage.delete_slicer(
                workbook, get_sheet(workbook, sheet), name, workspace.limits.max_cells
            )
        return Changed(sheet=sheet, name=name)
