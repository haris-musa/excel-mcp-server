"""Tools for defined names and cell notes."""

from typing import Annotated

from pydantic import Field

from excel_mcp.operations import names, notes
from excel_mcp.server.params import CellRef, SheetName, WorkbookPath
from excel_mcp.server.registry import ToolRegistry
from excel_mcp.workspace import Workspace, get_sheet

DefinedName = Annotated[
    str, Field(description="Letters, digits, underscores and periods, e.g. 'TaxRate'.")
]
NameScope = Annotated[
    str | None,
    Field(description="Sheet the name belongs to. Default: the whole workbook."),
]


def register(tools: ToolRegistry, workspace: Workspace) -> None:
    @tools.destroyer("Set defined name")
    def set_defined_name(
        path: WorkbookPath,
        name: DefinedName,
        refers_to: Annotated[
            str,
            Field(
                description="A range with its sheet, e.g. 'Data!$B$2:$B$100', or a constant "
                "such as '0.075'."
            ),
        ],
        sheet: NameScope = None,
    ) -> str:
        """Create a defined name for a range or constant, replacing a name of the same scope.

        Formulas can then use it, e.g. '=SUM(Sales)'. The reference follows the same safety
        rules as formulas. Names are listed by describe_workbook.
        """
        with workspace.edit(path) as workbook:
            replaced = names.set_defined_name(
                workbook, name, refers_to, get_sheet(workbook, sheet) if sheet else None
            )
        return f"{'Updated' if replaced else 'Created'} name {name!r} as {refers_to}."

    @tools.destroyer("Delete defined name")
    def delete_defined_name(path: WorkbookPath, name: DefinedName, sheet: NameScope = None) -> str:
        """Delete a defined name. Formulas that use it are not changed and will show #NAME?."""
        with workspace.edit(path) as workbook:
            names.delete_defined_name(workbook, name, get_sheet(workbook, sheet) if sheet else None)
        return f"Deleted name {name!r}."

    @tools.destroyer("Set note")
    def set_note(
        path: WorkbookPath,
        sheet: SheetName,
        cell: CellRef,
        text: Annotated[str, Field(min_length=1, description="The note's text.")],
        author: Annotated[str, Field(description="Name shown as the note's author.")] = "Claude",
    ) -> str:
        """Add a note to a cell, replacing the cell's existing note.

        Notes are listed by describe_sheet and shown when hovering over the cell in Excel.
        """
        with workspace.edit(path) as workbook:
            notes.set_note(get_sheet(workbook, sheet), cell, text, author)
        return f"Set the note on {sheet}!{cell}."

    @tools.destroyer("Delete note")
    def delete_note(path: WorkbookPath, sheet: SheetName, cell: CellRef) -> str:
        """Remove the note from a cell."""
        with workspace.edit(path) as workbook:
            notes.delete_note(get_sheet(workbook, sheet), cell)
        return f"Deleted the note on {sheet}!{cell}."
