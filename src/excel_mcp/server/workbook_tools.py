"""Tools for whole workbooks: create, describe, list, export and import."""

from typing import Annotated
from urllib.parse import quote

from mcp.types import BlobResourceContents, CallToolResult, EmbeddedResource, TextContent
from pydantic import Field

from excel_mcp.operations import files, inspect
from excel_mcp.operations.inspect import WorkbookFile, WorkbookInfo
from excel_mcp.operations.sheets import validate_sheet_name
from excel_mcp.server.params import WorkbookPath
from excel_mcp.server.registry import ToolRegistry
from excel_mcp.workspace import Workspace

XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def register(tools: ToolRegistry, workspace: Workspace) -> None:
    @tools.destroyer("Create workbook")
    def create_workbook(
        path: WorkbookPath,
        sheets: Annotated[
            list[str] | None,
            Field(min_length=1, description="Worksheet names, in order. Default: ['Sheet1']."),
        ] = None,
        overwrite: Annotated[
            bool, Field(description="Replace the file if it already exists.")
        ] = False,
    ) -> str:
        """Create a new, empty Excel workbook."""
        sheets = sheets or ["Sheet1"]
        for index, name in enumerate(sheets):
            validate_sheet_name(name, sheets[:index])
        created = workspace.create(path, sheets, overwrite=overwrite)
        return f"Created {workspace.display(created)} with sheets {sheets}."

    @tools.reader("Describe workbook")
    def describe_workbook(path: WorkbookPath) -> WorkbookInfo:
        """List a workbook's sheets with their used ranges, plus its defined names.

        Start here to learn a workbook's structure before reading or editing it.
        """
        resolved = workspace.resolve_existing(path)
        with workspace.read(path) as workbook:
            return inspect.describe_workbook(
                workbook, workspace.display(resolved), resolved.stat().st_size
            )

    @tools.reader("List workbooks")
    def list_workbooks(
        directory: Annotated[
            str,
            Field(
                description="Directory to search. Leave empty for the server's workbook "
                "directory; otherwise use an absolute path."
            ),
        ] = "",
        recursive: Annotated[bool, Field(description="Also search subdirectories.")] = False,
    ) -> list[WorkbookFile]:
        """List Excel files in a directory."""
        folder = workspace.resolve_directory(directory)
        return [
            WorkbookFile(path=workspace.display(found), size_bytes=found.stat().st_size)
            for found in inspect.list_workbooks(folder, recursive, limit=500)
            if workspace.paths.allows(found.resolve())
        ]

    @tools.reader("Export workbook")
    def export_workbook(path: WorkbookPath) -> CallToolResult:
        """Return the workbook file itself as an embedded base64 resource.

        Use this to hand a workbook to the user when the server runs remotely.
        """
        resolved = workspace.resolve_existing(path)
        name = workspace.display(resolved)
        return CallToolResult(
            content=[
                TextContent(type="text", text=f"Exported {name}."),
                EmbeddedResource(
                    type="resource",
                    resource=BlobResourceContents(
                        uri=f"excel-mcp://workbook/{quote(name)}",
                        mime_type=XLSX_MIME,
                        blob=files.encode_file(resolved),
                    ),
                ),
            ]
        )

    @tools.destroyer("Import workbook")
    def import_workbook(
        path: WorkbookPath,
        content_base64: Annotated[str, Field(description="The workbook file, base64 encoded.")],
        overwrite: Annotated[
            bool, Field(description="Replace the file if it already exists.")
        ] = False,
    ) -> str:
        """Save an uploaded workbook file on the server, e.g. to edit it remotely."""
        content = files.decode_workbook(content_base64, workspace.limits.max_file_bytes)
        stored = workspace.store(path, content, overwrite=overwrite)
        return f"Saved {workspace.display(stored)} ({len(content):,} bytes)."
