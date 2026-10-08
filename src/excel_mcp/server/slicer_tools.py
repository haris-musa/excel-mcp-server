"""Tools for slicers and timelines."""

from typing import Annotated, Literal

from pydantic import Field

from excel_mcp.operations import slicer_manage, slicers
from excel_mcp.operations.calculated import formula_values
from excel_mcp.operations.slicers import SlicerRequest, Source, Timeline
from excel_mcp.server.params import CellRef, SheetName, WorkbookPath
from excel_mcp.server.registry import ToolRegistry
from excel_mcp.workspace import Workspace, get_sheet


def register(tools: ToolRegistry, workspace: Workspace) -> None:
    @tools.writer("Add slicer")
    def add_slicer(
        path: WorkbookPath,
        sheet: Annotated[SheetName, Field(description="Sheet to put the slicer on.")],
        source: Annotated[Source, Field(description="The table or PivotTable to filter.")],
        field: Annotated[str, Field(description="Header of the column or field to filter by.")],
        cell: Annotated[CellRef, Field(description="Top-left cell of the slicer.")],
        width_cm: Annotated[
            float | None, Field(gt=0, le=100, description="Default: 5.1 (timeline: 9.3).")
        ] = None,
        height_cm: Annotated[
            float | None, Field(gt=0, le=100, description="Default: 7.4 (timeline: 3.8).")
        ] = None,
        caption: Annotated[
            str | None, Field(description="Header text. Default: the field.")
        ] = None,
        name: Annotated[str | None, Field(description="Default: the field name.")] = None,
        columns: Annotated[int, Field(ge=1, le=20, description="Columns of buttons.")] = 1,
        style: Annotated[
            str | None,
            Field(
                pattern=r"^(Time)?SlicerStyle(Light|Dark|Other)\d$",
                description="Built-in style, e.g. SlicerStyleDark2 (timeline: "
                "TimeSlicerStyleLight1). Default: Excel's.",
            ),
        ] = None,
        selected_items: Annotated[
            list[str], Field(max_length=10_000, description="Items to show. Default: all.")
        ] = [],  # noqa: B006
        sort: Annotated[
            Literal["ascending", "descending"], Field(description="Order of the items.")
        ] = "ascending",
        hide_empty_items: Annotated[
            bool, Field(description="Hide items that have no data.")
        ] = False,
        connect: Annotated[
            list[Source],
            Field(
                max_length=50,
                description="More PivotTables that share the source's data cache, as copies of "
                "a sheet do; the slicer filters them all.",
            ),
        ] = [],  # noqa: B006
        timeline: Annotated[
            Timeline | None,
            Field(
                description="Make a timeline instead of a slicer, for a date field of a "
                "PivotTable: time scale and the period shown."
            ),
        ] = None,
    ) -> str:
        """Add a slicer (Insert > Slicer) or timeline to a sheet, to filter a PivotTable or table.

        `selected_items` limits the data as clicking the buttons does: PivotTable items are hidden
        and its figures recalculated from the source; table rows are filtered and hidden.
        A PivotTable Excel refreshed after this server made it cannot be limited here; selecting
        every item works on any. describe_sheet lists slicers; delete_slicer removes one.
        """
        request = SlicerRequest(
            source, connect, field, cell, width_cm, height_cm, caption, name, columns, style,
            selected_items, sort, hide_empty_items, timeline,
        )  # fmt: skip
        with workspace.edit(path) as workbook:

            def results(owner, area):
                return formula_values(workspace, path, owner, area)

            made = slicers.add_slicer(
                workbook, get_sheet(workbook, sheet), request, results, workspace.limits.max_cells
            )
        return f"Added {'timeline' if timeline else 'slicer'} {made!r} to {sheet}!{cell}."

    @tools.destroyer("Delete slicer")
    def delete_slicer(
        path: WorkbookPath,
        sheet: SheetName,
        name: Annotated[str, Field(description="Slicer or timeline name, from describe_sheet.")],
    ) -> str:
        """Remove a slicer or timeline.

        As in Excel, a table and a PivotTable row, column or filter field stay filtered; a
        timeline's period and the hidden items of a field the PivotTable does not show are cleared.
        """
        with workspace.edit(path) as workbook:
            field = slicer_manage.delete_slicer(
                workbook, get_sheet(workbook, sheet), name, workspace.limits.max_cells
            )
        return f"Deleted {name!r} (field {field!r}) from {sheet}."
