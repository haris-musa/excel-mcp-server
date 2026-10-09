"""Tools for defined names and cell notes."""

from typing import Annotated

from pydantic import Field

from excel_mcp.operations import names, notes
from excel_mcp.server.params import CellRef, SheetName, WorkbookPath
from excel_mcp.server.registry import ToolRegistry
from excel_mcp.server.results import Changed
from excel_mcp.workspace import Workspace, get_sheet

DefinedName = Annotated[
    str, Field(description="Letters, digits, underscores and periods, e.g. 'TaxRate'.")
]
NameScope = Annotated[
    str | None,
    Field(description="Scope sheet. Default: the workbook."),
]


def register(tools: ToolRegistry, workspace: Workspace) -> None:
    @tools.destroyer("Set defined name")
    def set_defined_name(
        path: WorkbookPath,
        name: DefinedName,
        refers_to: Annotated[
            str,
            Field(description="Range with sheet, 'Data!$B$2:$B$100', or a constant, '0.075'."),
        ],
        sheet: NameScope = None,
    ) -> Changed:
        """Create a defined name for a range or constant (usable in formulas, '=SUM(Sales)'),
        replacing one of the same scope. The reference follows the formula safety rules."""
        with workspace.edit(path) as workbook:
            replaced = names.set_defined_name(
                workbook, name, refers_to, get_sheet(workbook, sheet) if sheet else None
            )
        note = "Replaced the earlier definition." if replaced else None
        return Changed(sheet=sheet, name=name, note=note)

    @tools.destroyer("Delete defined name")
    def delete_defined_name(
        path: WorkbookPath, name: DefinedName, sheet: NameScope = None
    ) -> Changed:
        """Delete a defined name; formulas using it will show #NAME?."""
        with workspace.edit(path) as workbook:
            names.delete_defined_name(workbook, name, get_sheet(workbook, sheet) if sheet else None)
        return Changed(sheet=sheet, name=name)

    @tools.destroyer("Set note")
    def set_note(
        path: WorkbookPath,
        sheet: SheetName,
        cell: CellRef,
        text: Annotated[str, Field(min_length=1)],
        author: str = "Claude",
    ) -> str:
        """Add a note to a cell, replacing its existing note."""
        with workspace.edit(path) as workbook:
            notes.set_note(get_sheet(workbook, sheet), cell, text, author)
        return f"Set the note on {sheet}!{cell}."

    @tools.destroyer("Delete note")
    def delete_note(path: WorkbookPath, sheet: SheetName, cell: CellRef) -> str:
        """Remove the note from a cell."""
        with workspace.edit(path) as workbook:
            notes.delete_note(get_sheet(workbook, sheet), cell)
        return f"Deleted the note on {sheet}!{cell}."
