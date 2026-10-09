"""Tools for whole workbooks: create, describe, list, export and import."""

from typing import Annotated
from urllib.parse import quote

from mcp.types import BlobResourceContents, CallToolResult, EmbeddedResource, TextContent
from pydantic import Field

from excel_mcp.operations import files, inspect, vba, workbook_settings
from excel_mcp.operations.inspect import WorkbookInfo
from excel_mcp.operations.sheet_counts import count_objects
from excel_mcp.operations.sheets import validate_sheet_name
from excel_mcp.operations.workbook_settings import WorkbookSettings
from excel_mcp.package.properties import read_company
from excel_mcp.server.params import WorkbookPath
from excel_mcp.server.registry import ToolRegistry
from excel_mcp.server.results import Changed
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
    ) -> Changed:
        """Create a new, empty Excel workbook."""
        sheets = sheets or ["Sheet1"]
        for index, name in enumerate(sheets):
            validate_sheet_name(name, sheets[:index])
        created = workspace.create(path, sheets, overwrite=overwrite)
        return Changed(path=workspace.display(created))

    @tools.reader("Describe workbook")
    def describe_workbook(path: WorkbookPath) -> WorkbookInfo:
        """List a workbook's sheets (visibility, used range, how many tables, charts,
        PivotTables, slicers and images each holds), defined names, properties and calculation
        settings.

        Start here. Reads each sheet once in full. Default and empty values are omitted;
        `has_vba` is only present when true, see read_vba.
        """
        with workspace.stream(path) as workbook:
            resolved = workspace.resolve(path)
            return inspect.describe_workbook(
                workbook, vba.has_vba(resolved), read_company(resolved), count_objects(resolved)
            )

    @tools.writer("Set workbook settings")
    def set_workbook_settings(path: WorkbookPath, settings: WorkbookSettings) -> Changed:
        """Set document properties (title, subject, author, keywords, company), calculation
        options (manual or automatic, iterative calculation, recalculation on load) and
        workbook structure protection.

        Structure protection discourages adding, deleting, renaming, moving and hiding sheets
        in Excel but is not security: it does not stop this server, and the password is weakly
        hashed.
        """
        with workspace.edit(path) as workbook:
            workbook_settings.apply_settings(workbook, settings)
        return Changed(path=path)

    @tools.reader("List workbooks")
    def list_workbooks(
        directory: Annotated[
            str,
            Field(description="Default: the server's workbook directory; else an absolute path."),
        ] = "",
        recursive: Annotated[bool, Field(description="Also search subdirectories.")] = False,
    ) -> dict[str, int]:
        """List Excel files in a directory as path to size in bytes."""
        folder = workspace.resolve_directory(directory)
        return {
            workspace.display(found): found.stat().st_size
            for found in inspect.list_workbooks(folder, recursive, limit=500)
            if workspace.paths.allows(found.resolve())
        }

    @tools.reader("Export workbook")
    def export_workbook(path: WorkbookPath) -> CallToolResult:
        """Return the workbook file as an embedded base64 resource, for remote servers."""
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
    ) -> Changed:
        """Save an uploaded workbook file on the server, e.g. to edit it remotely."""
        content = files.decode_workbook(content_base64, workspace.limits.max_file_bytes)
        stored = workspace.store(path, content, overwrite=overwrite)
        return Changed(path=workspace.display(stored))
