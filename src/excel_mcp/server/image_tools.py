"""Tools for pictures on a sheet."""

from typing import Annotated

from pydantic import Field

from excel_mcp.operations import images
from excel_mcp.server.params import CellRef, SheetName, WorkbookPath
from excel_mcp.server.registry import ToolRegistry
from excel_mcp.workspace import Workspace, get_sheet


def register(tools: ToolRegistry, workspace: Workspace) -> None:
    @tools.writer("Insert image")
    def insert_image(
        path: WorkbookPath,
        sheet: SheetName,
        image_path: Annotated[str, Field(description="PNG or JPEG file, in a workbook folder.")],
        cell: CellRef,
        width_cm: Annotated[
            float | None, Field(gt=0, le=100, description="Default: natural size.")
        ] = None,
        height_cm: Annotated[
            float | None, Field(gt=0, le=100, description="Alone, the ratio is kept.")
        ] = None,
    ) -> str:
        """Place a picture with its top-left corner at a cell.

        Give width_cm or height_cm to resize it keeping the ratio; give both to stretch it.
        describe_sheet lists images; delete_image removes one.
        """
        data = workspace.read_image(image_path)
        with workspace.edit(path) as workbook:
            images.insert_image(get_sheet(workbook, sheet), data, cell, width_cm, height_cm)
        return f"Inserted the image at {sheet}!{cell}."

    @tools.destroyer("Delete image")
    def delete_image(
        path: WorkbookPath,
        sheet: SheetName,
        index: Annotated[
            int,
            Field(ge=1, description="Image number from describe_sheet."),
        ],
    ) -> str:
        """Remove a picture from a sheet.

        Later images move up one index; call describe_sheet again before deleting another.
        """
        with workspace.edit(path) as workbook:
            removed = images.delete_image(get_sheet(workbook, sheet), index)
        return f"Deleted image {index} (at {removed.anchor}) from {sheet}."
