"""Cell notes (the yellow comment boxes shown when hovering over a cell)."""

from openpyxl.comments import Comment
from openpyxl.worksheet.worksheet import Worksheet
from pydantic import BaseModel

from excel_mcp.errors import InvalidArgumentError
from excel_mcp.operations.cells import writable_cell
from excel_mcp.refs import parse_cell


class NoteInfo(BaseModel):
    cell: str
    text: str
    author: str


def list_notes(sheet: Worksheet) -> list[NoteInfo]:
    return [
        NoteInfo(cell=cell.coordinate, text=cell.comment.text, author=cell.comment.author or "")
        for _, cell in sorted(sheet._cells.items())
        if cell.comment and not has_thread(cell.comment)
    ]


def set_note(sheet: Worksheet, ref: str, text: str, author: str) -> None:
    if not text.strip():
        raise InvalidArgumentError("Note text cannot be empty; use delete_note to remove a note.")
    cell = writable_cell(sheet, *parse_cell(ref))
    if has_thread(cell.comment):
        raise InvalidArgumentError(
            f"{cell.coordinate} has a threaded comment; delete_note removes it first."
        )
    cell.comment = new_note(text, author)


def new_note(text: str, author: str) -> Comment:
    # openpyxl sizes every note box 144x79 points, which hides most of a long note.
    lines = text.count("\n") + len(text) // 40 + 2
    return Comment(text, author, height=max(79, 15 * lines), width=300)


def has_thread(comment: Comment | None) -> bool:
    """Whether a note is the stand-in for a threaded comment."""
    return comment is not None and (comment.author or "").startswith("tc=")


def delete_note(sheet: Worksheet, ref: str) -> None:
    cell = writable_cell(sheet, *parse_cell(ref))
    if cell.comment is None:
        raise InvalidArgumentError(f"{cell.coordinate} has no note.")
    cell.comment = None
