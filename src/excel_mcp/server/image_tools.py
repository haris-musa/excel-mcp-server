"""Tools for pictures on a sheet."""

from typing import Annotated

from pydantic import Field

from excel_mcp.operations import images
from excel_mcp.server.params import CellRef, SheetName, WorkbookPath
from excel_mcp.server.registry import ToolRegistry
from excel_mcp.server.results import Changed
from excel_mcp.workspace import Workspace, get_sheet


def register(tools: ToolRegistry, workspace: Workspace) -> None:
    @tools.writer("Insert image")
    def insert_image(
        path: WorkbookPath,
        sheet: SheetName,
        image_path: Annotated[str, Field(description="PNG or JPEG in a workbook folder.")],
        at: CellRef,
        width_cm: Annotated[
            float | None, Field(gt=0, le=100, description="Default: natural size.")
        ] = None,
        height_cm: Annotated[
            float | None, Field(gt=0, le=100, description="Alone, the ratio is kept.")
        ] = None,
        name: Annotated[
            str | None, Field(description="Unique among the sheet's charts and images.")
        ] = None,
    ) -> Changed:
        """Place an image at a cell. Returns its name and cells. One of width_cm or height_cm
        keeps the ratio; both stretch it."""
        data = workspace.read_image(image_path)
        with workspace.edit(path) as workbook:
            placed = images.insert_image(
                get_sheet(workbook, sheet), data, at, width_cm, height_cm, name
            )
        return Changed(sheet=sheet, name=placed.name, range=placed.range)

    @tools.destroyer("Delete image")
    def delete_image(
        path: WorkbookPath,
        sheet: SheetName,
        name: Annotated[str, Field(description="Name from describe_sheet.")],
    ) -> Changed:
        """Remove an image from a sheet."""
        with workspace.edit(path) as workbook:
            removed = images.delete_image(get_sheet(workbook, sheet), name)
        return Changed(sheet=sheet, name=removed.name, range=removed.range)
