"""Threaded comments (Review > New Comment): a conversation on a cell, with replies.

Excel stores one as a ``threadedComment`` entry per comment, a ``person`` for each author,
and a note on the cell that older versions of Excel show instead. A cell holds a note or a
thread, never both. The parts are kept by the package layer; this module edits their text.
"""

import re
import uuid
from datetime import UTC, datetime
from typing import cast
from xml.sax.saxutils import escape, quoteattr

from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet
from pydantic import BaseModel

from excel_mcp.errors import InvalidArgumentError
from excel_mcp.operations.cells import writable_cell
from excel_mcp.operations.notes import delete_note, has_thread, new_note
from excel_mcp.package import Link, Part, state_of
from excel_mcp.package.consistency import THREADED_COMMENTS
from excel_mcp.package.scan import attributes, scan, unescape
from excel_mcp.refs import parse_cell

NAMESPACE = "http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments"
_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_REL = "http://schemas.microsoft.com/office/2017/10/relationships/"
_PERSONS = "application/vnd.ms-excel.person+xml"
_XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
_ROOT = f'<ThreadedComments xmlns="{NAMESPACE}" xmlns:x="{_MAIN}">'
_PERSON_LIST = f'<personList xmlns="{NAMESPACE}" xmlns:x="{_MAIN}">'
_FALLBACK = (
    "[Threaded comment]\n\nYour version of Excel allows you to read this threaded comment; "
    "however, any edits to it will get removed if the file is opened in a newer version of "
    "Excel. Learn more: https://go.microsoft.com/fwlink/?linkid=870924\n\nComment:\n    {}"
)
_TEXT = re.compile(r"<text>(.*?)</text>", re.S)
_PERSON = re.compile(r"<person\b[^>]*/>")


class CommentInfo(BaseModel):
    author: str
    date: str
    text: str


class ThreadInfo(BaseModel):
    cell: str
    resolved: bool = False
    comments: list[CommentInfo]


class _Entry:
    """One ``threadedComment`` element."""

    def __init__(self, raw: str) -> None:
        self.raw = raw
        self.values = attributes(raw[: raw.index(">") + 1])
        match = _TEXT.search(raw)
        self.text = unescape(match[1]) if match else ""

    @property
    def is_reply(self) -> bool:
        return "parentId" in self.values


def _new_id() -> str:
    return "{" + str(uuid.uuid4()).upper() + "}"


def _timestamp() -> str:
    now = datetime.now(UTC)
    return f"{now:%Y-%m-%dT%H:%M:%S}.{now.microsecond // 10000:02d}"


def _entry(ref: str, person: str, text: str, parent: str | None) -> _Entry:
    ids = f'id="{_new_id()}"' + (f' parentId="{parent}"' if parent else "")
    return _Entry(
        f'<threadedComment ref="{ref}" dT="{_timestamp()}" personId="{person}" {ids} done="0">'
        f"<text>{escape(text)}</text></threadedComment>"
    )


def _find(links: list[Link], kind: str) -> Link | None:
    return next((x for x in links if x.type == kind and isinstance(x.target, Part)), None)


class _Sheet:
    """The threaded comments of one sheet and the persons of its workbook."""

    def __init__(self, sheet: Worksheet) -> None:
        self.sheet = sheet
        self.state = state_of(cast(Workbook, sheet.parent))
        self.package = self.state.sheet(sheet)
        self.link = _find(self.package.links, _REL + "threadedComment")
        self.entries: list[_Entry] = []
        if self.link and isinstance(self.link.target, Part):
            document = scan(self.link.target.data)
            self.entries = [_Entry(document.raw(child)) for child in document.children]

    @property
    def persons(self) -> Part | None:
        link = _find(self.state.workbook.links, _REL + "person")
        return link.target if link and isinstance(link.target, Part) else None

    def thread(self, ref: str) -> list[_Entry]:
        return [e for e in self.entries if e.values["ref"] == ref]

    def names(self) -> dict[str, str]:
        persons = self.persons
        tags = _PERSON.findall(persons.data.decode() if persons else "")
        return {p["id"]: p.get("displayName", "") for p in map(attributes, tags)}

    def person(self, author: str) -> str:
        """The id of the person with this name, who is added if the file has none."""
        for identifier, name in self.names().items():
            if name == author:
                return identifier
        identifier = _new_id()
        tag = (
            f"<person displayName={quoteattr(author)} id={quoteattr(identifier)} "
            f'userId={quoteattr(author)} providerId="None"/>'
        )
        if persons := self.persons:
            persons.data = persons.data.replace(b"</personList>", f"{tag}</personList>".encode())
        else:
            data = f"{_XML}{_PERSON_LIST}{tag}</personList>".encode()
            part = Part("xl/persons/person.xml", _PERSONS, data)
            self.state.workbook.links.append(Link("rIdPersons", _REL + "person", part))
        return identifier

    def save(self) -> None:
        """Write the entries back and show each thread's note on its cell."""
        data = f"{_XML}{_ROOT}{''.join(e.raw for e in self.entries)}</ThreadedComments>".encode()
        if not self.entries:
            if self.link:
                self.package.links.remove(self.link)
        elif self.link and isinstance(self.link.target, Part):
            self.link.target.data = data
        else:
            part = Part("xl/threadedComments/threadedComment1.xml", THREADED_COMMENTS, data)
            self.link = Link("rIdThreads", _REL + "threadedComment", part)
            self.package.links.append(self.link)
        for root in self.entries:
            if not root.is_reply:
                thread = self.thread(root.values["ref"])
                replies = "".join(f"\nReply:\n    {r.text}" for r in thread[1:])
                self.sheet[root.values["ref"]].comment = new_note(
                    _FALLBACK.format(root.text) + replies, "tc=" + root.values["id"]
                )


def add_comment(sheet: Worksheet, ref: str, text: str, author: str) -> bool:
    """Start a thread on a cell, or reply to its thread; returns whether it was a reply."""
    cell = writable_cell(sheet, *parse_cell(ref))
    if cell.comment and not has_thread(cell.comment):
        raise InvalidArgumentError(
            f"{cell.coordinate} has a note; delete_note removes it before a thread can start."
        )
    comments = _Sheet(sheet)
    thread = comments.thread(cell.coordinate)
    person = comments.person(author)
    parent = thread[0].values["id"] if thread else None
    comments.entries.append(_entry(cell.coordinate, person, text, parent))
    comments.save()
    return bool(thread)


def delete_comment(sheet: Worksheet, ref: str, reply: int) -> None:
    """Delete a cell's note or thread (``reply`` 0), or reply number ``reply`` of its thread."""
    cell = writable_cell(sheet, *parse_cell(ref))
    if not has_thread(cell.comment):
        if reply:
            raise InvalidArgumentError(f"{cell.coordinate} has no threaded comment.")
        delete_note(sheet, ref)
        return
    comments = _Sheet(sheet)
    thread = _thread(comments, cell.coordinate)
    if reply > len(thread) - 1:
        raise InvalidArgumentError(
            f"{cell.coordinate} has {len(thread) - 1} replies; reply {reply} does not exist."
        )
    gone = thread if reply == 0 else [thread[reply]]
    comments.entries = [e for e in comments.entries if e not in gone]
    if reply == 0:
        cell.comment = None
    comments.save()


def resolve_thread(sheet: Worksheet, ref: str, resolved: bool) -> None:
    cell = writable_cell(sheet, *parse_cell(ref))
    comments = _Sheet(sheet)
    root = _thread(comments, cell.coordinate)[0]
    start = re.sub(r'\s+done="[^"]*"', "", root.raw[: root.raw.index(">")])
    updated = f'{start} done="{int(resolved)}"{root.raw[root.raw.index(">") :]}'
    comments.entries[comments.entries.index(root)] = _Entry(updated)
    comments.save()


def _thread(comments: _Sheet, ref: str) -> list[_Entry]:
    thread = comments.thread(ref)
    if not thread:
        raise InvalidArgumentError(f"{ref} has no threaded comment.")
    return thread


def list_threads(sheet: Worksheet) -> list[ThreadInfo]:
    comments = _Sheet(sheet)
    names = comments.names()
    return [
        ThreadInfo(
            cell=root.values["ref"],
            resolved=root.values.get("done") in ("1", "true"),
            comments=[
                CommentInfo(
                    author=names.get(e.values["personId"], ""), date=e.values["dT"], text=e.text
                )
                for e in comments.thread(root.values["ref"])
            ],
        )
        for root in comments.entries
        if not root.is_reply
    ]


def copy_threads(source: Worksheet, target: Worksheet) -> None:
    """Give a copy of a sheet its own threads, with new ids as Excel makes them."""
    original = _Sheet(source)
    if not original.entries:
        return
    renamed = {e.values["id"]: _new_id() for e in original.entries}
    copy = _Sheet(target)
    for entry in original.entries:
        raw = entry.raw
        for old, new in renamed.items():
            raw = raw.replace(old, new)
        copy.entries.append(_Entry(raw))
    copy.save()
