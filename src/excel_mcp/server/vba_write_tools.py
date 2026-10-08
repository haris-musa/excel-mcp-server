"""Tools that write VBA macros. Registered only when the server allows VBA writing."""

from typing import Annotated, Literal

from pydantic import Field

from excel_mcp.operations import vba_write
from excel_mcp.operations.vba_write import VbaChange
from excel_mcp.server.params import WorkbookPath
from excel_mcp.server.registry import ToolRegistry
from excel_mcp.workspace import Workspace


def register(tools: ToolRegistry, workspace: Workspace) -> None:
    @tools.destroyer("Write VBA module")
    def write_vba_module(
        path: WorkbookPath,
        module: Annotated[
            str,
            Field(
                description="Module name, e.g. 'Module1', or 'ThisWorkbook' / 'Sheet1' to set "
                "the code behind the workbook or a sheet (for event handlers)."
            ),
        ],
        code: Annotated[
            str,
            Field(
                max_length=vba_write.MAX_CODE_CHARS,
                description="The module's complete VBA code; it replaces the existing code. "
                "Empty clears the module. No 'Attribute' lines.",
            ),
        ],
        kind: Annotated[
            Literal["standard", "class"],
            Field(description="Type of the module to create; ignored if it already exists."),
        ] = "standard",
    ) -> VbaChange:
        """Set the VBA code of a module in an .xlsm or .xltm workbook, creating it if new.

        Existing modules keep their type; other modules and forms are untouched. A macro
        workbook without a VBA project gets one. .xlsx files cannot hold macros, and digitally
        signed projects are refused (changing them would invalidate the signature). The code is
        only stored, never run here: tell the user to review it before enabling macros in Excel.
        """
        with workspace.edit(path) as workbook:
            return vba_write.write_module(workbook, module, code, kind)

    @tools.destroyer("Delete VBA module")
    def delete_vba_module(
        path: WorkbookPath,
        module: Annotated[str, Field(description="Name of a standard or class module.")],
    ) -> VbaChange:
        """Delete a standard or class module from an .xlsm or .xltm workbook.

        Workbook and sheet modules cannot be deleted; clear them with write_vba_module.
        """
        with workspace.edit(path) as workbook:
            return vba_write.delete_module(workbook, module)
