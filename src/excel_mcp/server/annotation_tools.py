"""Tools for defined names and cell notes."""

from typing import Annotated

from pydantic import Field

from excel_mcp.operations import names, notes, threads
from excel_mcp.server.params import CellRef, SheetName, WorkbookPath
from excel_mcp.server.registry import ToolRegistry
from excel_mcp.workspace import Workspace, get_sheet

DefinedName = Annotated[
    str, Field(description="Letters, digits, underscores and periods, e.g. 'TaxRate'.")
]
NameScope = Annotated[
    str | None,
    Field(description="Sheet the name is scoped to. Default: the workbook."),
]


def register(tools: ToolRegistry, workspace: Workspace) -> None:
    @tools.destroyer("Set defined name")
    def set_defined_name(
        path: WorkbookPath,
        name: DefinedName,
        refers_to: Annotated[
            str,
            Field(description="Range with sheet, e.g. 'Data!$B$2:$B$100', or a constant: '0.075'."),
        ],
        sheet: NameScope = None,
    ) -> str:
        """Create a defined name for a range or constant, replacing a name of the same scope.

        Formulas can then use it, e.g. '=SUM(Sales)'. The reference follows the formula
        safety rules. describe_workbook lists names.
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
        text: Annotated[str, Field(min_length=1, description="Note or comment text.")],
        author: Annotated[str, Field(description="Shown as the author.")] = "Claude",
        threaded: Annotated[
            bool,
            Field(
                description="true: a threaded comment (Review > New Comment), a reply if the "
                "cell has a thread. false: a note, replacing the cell's note."
            ),
        ] = False,
    ) -> str:
        """Add a note, or a threaded comment, to a cell. A cell has one or the other.

        describe_sheet lists both. Excel shows notes on hover; resolve_comment resolves threads.
        """
        with workspace.edit(path) as workbook:
            target = get_sheet(workbook, sheet)
            if not threaded:
                notes.set_note(target, cell, text, author)
                return f"Set the note on {sheet}!{cell}."
            reply = threads.add_comment(target, cell, text, author)
        return f"Added {'a reply' if reply else 'a thread'} on {sheet}!{cell}."

    @tools.destroyer("Delete note")
    def delete_note(
        path: WorkbookPath,
        sheet: SheetName,
        cell: CellRef,
        reply: Annotated[
            int,
            Field(ge=0, description="Reply number in the cell's thread (1 = first). 0: all."),
        ] = 0,
    ) -> str:
        """Remove the note or whole threaded comment from a cell, or one reply of a thread."""
        with workspace.edit(path) as workbook:
            threads.delete_comment(get_sheet(workbook, sheet), cell, reply)
        return f"Deleted the {'reply' if reply else 'note or thread'} on {sheet}!{cell}."

    @tools.writer("Resolve comment")
    def resolve_comment(
        path: WorkbookPath,
        sheet: SheetName,
        cell: CellRef,
        resolved: Annotated[bool, Field(description="false reopens the thread.")] = True,
    ) -> str:
        """Resolve or reopen the threaded comment on a cell."""
        with workspace.edit(path) as workbook:
            threads.resolve_thread(get_sheet(workbook, sheet), cell, resolved)
        return f"{'Resolved' if resolved else 'Reopened'} the thread on {sheet}!{cell}."
