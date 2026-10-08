"""The VML drawing of a sheet: note boxes, which openpyxl writes, and form controls.

openpyxl rebuilds the VML for the notes of a sheet and drops every other shape in it, among
them the shapes of form controls and OLE objects. Those are kept and merged back.
"""

import re

from excel_mcp.package.patch import Edit, splice
from excel_mcp.package.scan import Scan, scan, with_namespaces

_NOTE_TYPE = "_x0000_t202"
_OBJECT_TYPE = re.compile(r'ObjectType="([^"]*)"')
_TYPE_REFERENCE = re.compile(r'\btype="#([^"]*)"')
_SHAPE_NUMBER = re.compile(r'(\bid="_x0000_s)(\d+)(")')
_ANY_NUMBER = re.compile(r"_x0000_s(\d+)")
_ID = re.compile(r'\bid="([^"]*)"')
_FIRST_ID = 1025


def preserved_shapes(document: Scan) -> list[str]:
    """What a VML drawing holds besides the note boxes: the layout, shape types and shapes."""
    shapes = []
    for child in document.children:
        kind = _OBJECT_TYPE.search(document.raw(child))
        if child.local == "shape" and not (kind and kind[1] == "Note"):
            shapes.append(document.raw(child))
    used = {m[1] for shape in shapes for m in _TYPE_REFERENCE.finditer(shape)} - {_NOTE_TYPE}
    kept = [
        document.raw(c)
        for c in document.children
        if c.local == "shapelayout" or (c.local == "shapetype" and _id_of(document.raw(c)) in used)
    ]
    return kept + shapes if shapes else []


def merge(written: bytes | None, preserved: list[str], namespaces: dict[str, str]) -> bytes:
    """The VML openpyxl wrote for the notes, with the preserved shapes added.

    Notes get ids above those of the preserved shapes, which other parts of the file name.
    """
    items = [with_namespaces(item, namespaces) for item in preserved]
    if written is None:
        return f"<xml>{''.join(items)}</xml>".encode()
    document = scan(written)
    present = {c.local for c in document.children}
    types = {_id_of(document.raw(c)) for c in document.children if c.local == "shapetype"}
    additions = [item for item in items if not _is_present(item, present, types)]
    taken = {int(n) for item in items for n in _ANY_NUMBER.findall(item)}
    edits = _renumbered(written, document, max(taken | {_FIRST_ID}) + 1)
    edits.append((document.close_tag, document.close_tag, "".join(additions).encode("utf-8")))
    return splice(written, edits)


def _is_present(item: str, present: set[str], types: set[str]) -> bool:
    """Whether the written VML has the layout or the shape type already."""
    if re.match(r"<[\w.-]*:?shapelayout\b", item):
        return "shapelayout" in present
    return bool(re.match(r"<[\w.-]*:?shapetype\b", item)) and _id_of(item) in types


def _id_of(xml: str) -> str:
    match = _ID.search(xml[: xml.index(">") + 1])
    return match[1] if match else ""


def _renumbered(data: bytes, document: Scan, first: int) -> list[Edit]:
    edits: list[Edit] = []
    for child in document.children:
        if child.local == "shape":
            tag = data[child.start : data.index(b">", child.start) + 1]
            text = _SHAPE_NUMBER.sub(rf"\g<1>{first}\g<3>", tag.decode("utf-8"), count=1)
            edits.append((child.start, child.start + len(tag), text.encode("utf-8")))
            first += 1
    return edits
