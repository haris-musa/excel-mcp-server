"""Editing the XML that openpyxl wrote, as text, so that nothing else in it changes."""

import re
from collections.abc import Callable
from itertools import pairwise

from excel_mcp.errors import WorkbookError
from excel_mcp.package.scan import Scan, scan

# The order of the children of a worksheet and of a workbook in the schema. Excel rejects
# a file whose elements are in another order.
SHEET_ORDER = [
    "sheetPr",
    "dimension",
    "sheetViews",
    "sheetFormatPr",
    "cols",
    "sheetData",
    "sheetCalcPr",
    "sheetProtection",
    "protectedRanges",
    "scenarios",
    "autoFilter",
    "sortState",
    "dataConsolidate",
    "customSheetViews",
    "mergeCells",
    "phoneticPr",
    "conditionalFormatting",
    "dataValidations",
    "hyperlinks",
    "printOptions",
    "pageMargins",
    "pageSetup",
    "headerFooter",
    "rowBreaks",
    "colBreaks",
    "customProperties",
    "cellWatches",
    "ignoredErrors",
    "smartTags",
    "drawing",
    "legacyDrawing",
    "legacyDrawingHF",
    "drawingHF",
    "picture",
    "oleObjects",
    "controls",
    "webPublishItems",
    "tableParts",
    "extLst",
]
WORKBOOK_ORDER = [
    "fileVersion",
    "fileSharing",
    "workbookPr",
    "workbookProtection",
    "bookViews",
    "sheets",
    "functionGroups",
    "externalReferences",
    "definedNames",
    "calcPr",
    "oleSize",
    "customWorkbookViews",
    "pivotCaches",
    "smartTagPr",
    "smartTagTypes",
    "webPublishing",
    "fileRecoveryPr",
    "webPublishObjects",
    "extLst",
]

Edit = tuple[int, int, bytes]
CellEdit = Callable[[str, str], tuple[str, str]]


def splice(data: bytes, edits: list[Edit]) -> bytes:
    """Replace the spans ``(start, end, text)``, which must not overlap."""
    result = bytearray()
    last = 0
    for start, end, text in sorted(edits, key=lambda edit: edit[:2]):
        result += data[last:start] + text
        last = end
    return bytes(result + data[last:])


def sheet_data_end(data: bytes) -> int:
    """Where the ``sheetData`` element of a worksheet ends."""
    closing = data.rfind(b"</sheetData>")
    if closing >= 0:
        return closing + len(b"</sheetData>")
    empty = re.search(rb"<sheetData\s*/>", data)
    if empty is None:
        raise WorkbookError("A worksheet has no sheetData.")
    return empty.end()


def insertions(document: Scan, new: list[tuple[str, str]], order: list[str]) -> list[Edit]:
    """Where to add elements, as (local name, XML), so that the schema order is kept."""
    rank = {name: position for position, name in enumerate(order)}
    grouped: dict[int, list[str]] = {}
    for name, xml in sorted(new, key=lambda item: rank[item[0]]):
        later = next((c for c in document.children if rank.get(c.local, 0) > rank[name]), None)
        grouped.setdefault(later.start if later else document.close_tag, []).append(xml)
    return [(point, point, "".join(xml).encode()) for point, xml in grouped.items()]


def append_to_root(data: bytes, xml: str) -> bytes:
    """Add an element as the last child of the root."""
    close = scan(data).close_tag
    return data[:close] + xml.encode() + data[close:]


def rule_extensions(
    data: bytes, document: Scan, extensions: dict[tuple[str, str], str]
) -> tuple[list[Edit], list[str]]:
    """Put the ``<extLst>`` of Excel's conditional formatting rules back into their rules.

    Returns the edits, and the extensions that found their rule.
    """
    edits: list[Edit] = []
    restored: list[str] = []
    for block in document.children:
        if block.local != "conditionalFormatting":
            continue
        text = data[block.start : block.end]
        sqref = re.search(rb'sqref="([^"]*)"', text[: text.index(b">") + 1])
        for rule in scan(text).children:
            rule_text = text[rule.start : rule.end]
            priority = re.search(rb'priority="([^"]*)"', rule_text[: rule_text.index(b">") + 1])
            key = (sqref[1].decode() if sqref else "", priority[1].decode() if priority else "")
            if key in extensions:
                offset = block.start + rule.start
                edits.append(_with_extension(rule_text, rule.qname, extensions[key], offset))
                restored.append(extensions[key])
    return edits, restored


def _with_extension(rule: bytes, qname: str, extension: str, offset: int) -> Edit:
    if rule.endswith(b"/>"):
        text = rule[:-2].rstrip() + b">" + extension.encode() + f"</{qname}>".encode()
    else:
        close = rule.rindex(b"</")
        text = rule[:close] + extension.encode() + rule[close:]
    return (offset, offset + len(rule), text)


# Looking each cell up by its address reads the sheet once per cell, which is fastest for a
# few cells; for more, one pass over all the cells does not grow with the number to patch.
_FEW_CELLS = 16
_CELL = re.compile(rb'<c r="([^"]*)"([^>]*?)(?:/>|>(.*?)</c>)', re.S)


def patch_cells(data: bytes, edits: dict[str, CellEdit]) -> bytes:
    """Rewrite chosen cells of ``sheetData``; an edit gets a cell's attributes and its content."""
    if not edits:
        return data
    if len(edits) <= _FEW_CELLS:
        return splice(data, _located_edits(data, edits))

    def rewrite(match: re.Match[bytes]) -> bytes:
        edit = edits.get(match[1].decode())
        return _rewritten(match, edit) if edit else match[0]

    return _CELL.sub(rewrite, data)


def _located_edits(data: bytes, edits: dict[str, CellEdit]) -> list[Edit]:
    found: list[Edit] = []
    for coordinate, edit in edits.items():
        opening = b'<c r="%b"' % coordinate.encode()
        start = data.find(opening)
        while start >= 0:
            match = _CELL.match(data, start)
            if match is None:
                raise WorkbookError(f"Cell {coordinate} of a worksheet is malformed.")
            found.append((match.start(), match.end(), _rewritten(match, edit)))
            start = data.find(opening, match.end())
    found.sort()
    if any(before[1] > after[0] for before, after in pairwise(found)):
        raise WorkbookError("A worksheet has cells inside other cells.")
    return found


def _rewritten(match: re.Match[bytes], edit: CellEdit) -> bytes:
    attributes, body = edit(match[2].decode(), (match[3] or b"").decode())
    start = f'<c r="{match[1].decode()}"{attributes}'
    return f"{start}>{body}</c>".encode() if body else f"{start}/>".encode()
