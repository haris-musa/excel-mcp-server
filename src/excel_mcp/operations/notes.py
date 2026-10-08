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
        if cell.comment
    ]


def set_note(sheet: Worksheet, ref: str, text: str, author: str) -> None:
    if not text.strip():
        raise InvalidArgumentError("Note text cannot be empty; use delete_note to remove a note.")
    cell = writable_cell(sheet, *parse_cell(ref))
    # openpyxl sizes every note box 144x79 points, which hides most of a long note.
    lines = text.count("\n") + len(text) // 40 + 2
    cell.comment = Comment(text, author, height=max(79, 15 * lines), width=300)


def delete_note(sheet: Worksheet, ref: str) -> None:
    cell = writable_cell(sheet, *parse_cell(ref))
    if cell.comment is None:
        raise InvalidArgumentError(f"{cell.coordinate} has no note.")
    cell.comment = None
