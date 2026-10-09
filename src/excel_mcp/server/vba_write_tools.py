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
            Field(description="e.g. 'Module1'; 'ThisWorkbook' or a sheet name for event handlers."),
        ],
        code: Annotated[
            str,
            Field(
                max_length=vba_write.MAX_CODE_CHARS,
                description="Complete VBA code, replacing the existing code; empty clears. "
                "No 'Attribute' lines.",
            ),
        ],
        kind: Annotated[
            Literal["standard", "class"],
            Field(description="For a new module; existing ones keep theirs."),
        ] = "standard",
    ) -> VbaChange:
        """Set the VBA code of a module in an .xlsm or .xltm workbook (created if new, with
        a VBA project if there is none). Signed projects are refused. The code is only stored,
        never run: tell the user to review it before enabling macros in Excel.
        """
        with workspace.edit(path) as workbook:
            return vba_write.write_module(workbook, module, code, kind)

    @tools.destroyer("Delete VBA module")
    def delete_vba_module(
        path: WorkbookPath,
        module: Annotated[str, Field(description="Name of a standard or class module.")],
    ) -> VbaChange:
        """Delete a standard or class module (workbook and sheet modules: clear them with
        write_vba_module)."""
        with workspace.edit(path) as workbook:
            return vba_write.delete_module(workbook, module)
