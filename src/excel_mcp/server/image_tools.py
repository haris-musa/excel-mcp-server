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
        image_path: Annotated[
            str, Field(description="PNG or JPEG file, in the same folders as workbooks.")
        ],
        cell: CellRef,
        width_cm: Annotated[
            float | None, Field(gt=0, le=100, description="Width in cm. Default: natural size.")
        ] = None,
        height_cm: Annotated[
            float | None,
            Field(gt=0, le=100, description="Height in cm; with only one size the ratio is kept."),
        ] = None,
    ) -> str:
        """Place a PNG or JPEG picture with its top-left corner at a cell.

        Give width_cm or height_cm to resize it; give both only to stretch it. List a
        sheet's images with describe_sheet and remove one with delete_image.
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
            Field(
                ge=1,
                description="1-based image number, as listed under 'images' by describe_sheet.",
            ),
        ],
    ) -> str:
        """Remove a picture from a sheet.

        Images after the removed one move up by one index, so call describe_sheet again
        before deleting another.
        """
        with workspace.edit(path) as workbook:
            removed = images.delete_image(get_sheet(workbook, sheet), index)
        return f"Deleted image {index} (at {removed.anchor}) from {sheet}."
