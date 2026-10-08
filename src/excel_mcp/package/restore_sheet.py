"""Putting what openpyxl dropped back into one worksheet."""

import re
from typing import TYPE_CHECKING
from xml.sax.saxutils import escape

from openpyxl.worksheet.formula import ArrayFormula
from openpyxl.worksheet.worksheet import Worksheet

from excel_mcp.package import consistency, pivots, vml
from excel_mcp.package import extensions as extensions_module
from excel_mcp.package.model import CellMark, Link, Part, SheetPackage
from excel_mcp.package.opc import REL_BASE, XML_DECLARATION, Rel, free_id
from excel_mcp.package.patch import (
    SHEET_ORDER,
    CellEdit,
    Edit,
    append_to_root,
    insertions,
    patch_cells,
    rule_extensions,
    sheet_data_end,
    splice,
)
from excel_mcp.package.scan import REL_NS, Scan, relationship_ids, scan, with_namespaces

if TYPE_CHECKING:
    from excel_mcp.package.restore import Restorer

_DRAWING_TYPE = "application/vnd.openxmlformats-officedocument.drawing+xml"
_VML_TYPE = "application/vnd.openxmlformats-officedocument.vmlDrawing"
_SHAPE_ID = re.compile(r'(<(?:[\w.-]+:)?cNvPr\b[^>]*?\bid=")(\d+)(")')
_RULE_ID = re.compile(r"<x14:id>(\{[^}]*\})</x14:id>")
_EMPTY_VALUE = re.compile(r"<v\s*/>|<v></v>")
_TYPE_ATTRIBUTE = re.compile(r'\st="[^"]*"')
_EMPTY_DRAWING = (
    f'{XML_DECLARATION}<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/'
    'spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
    "</xdr:wsDr>"
)


def restore_sheet(restorer: "Restorer", sheet: Worksheet, package: SheetPackage, part: str) -> None:
    """Merge the preserved content of a sheet into the XML openpyxl wrote for it."""
    data = restorer.package.read(part)
    created = _drawing(restorer, part, package)
    document = scan(data, "sheetData")
    created_vml = _vml(restorer, part, package, document)
    edits: list[Edit] = []
    restored: list[str] = []
    if package.rule_extensions:
        edits, restored = rule_extensions(data, document, package.rule_extensions)
    extensions = _extensions(restorer, package, restored)
    links = _threaded_comments(sheet, package.links)
    names = [name for name, _ in package.elements]
    texts = [xml for _, xml in package.elements] + list(extensions.values())
    texts = restorer.attach(part, links, texts, package.namespaces)
    new = list(zip(names, texts[: len(names)], strict=True))
    if created:
        new.append(("drawing", created))
    if created_vml:
        new.append(("legacyDrawing", created_vml))
    if extensions:
        new.append(("extLst", f"<extLst>{''.join(texts[len(names) :])}</extLst>"))
    _pivot_extensions(restorer, part, package)
    end = sheet_data_end(data)
    edits += insertions(document, new, SHEET_ORDER)
    head = patch_cells(data[:end], _cell_edits(restorer, sheet, package))
    restorer.package.write(part, head + splice(data[end:], _shifted(edits, end)))


def _extensions(restorer: "Restorer", package: SheetPackage, restored: list[str]) -> dict[str, str]:
    """The sheet's extensions, without what points at things that are gone."""
    extensions = dict(package.extensions)
    dead = {i for xml in package.rule_extensions.values() for i in _RULE_ID.findall(xml)}
    dead -= {i for xml in restored for i in _RULE_ID.findall(xml)}
    extensions_module.without_rules(extensions, dead)
    extensions_module.forget_sheets(extensions, restorer.gone)
    return extensions


def _threaded_comments(sheet: Worksheet, links: list[Link]) -> list[Link]:
    """Point threaded comments at the notes they are shown with; drop those without a note."""
    live = {
        cell.comment.author[3:]: cell.coordinate
        for cell in sheet._cells.values()
        if cell.comment and (cell.comment.author or "").startswith("tc=")
    }
    kept = []
    for link in links:
        target = link.target
        if isinstance(target, Part) and target.content_type == consistency.THREADED_COMMENTS:
            updated = consistency.reconcile_threads(target, live)
            if updated is None:
                continue
            link = Link(link.id, link.type, updated)
        kept.append(link)
    return kept


def _cell_edits(
    restorer: "Restorer", sheet: Worksheet, package: SheetPackage
) -> dict[str, CellEdit]:
    edits = {}
    for mark in package.marks:
        cell = mark.cell
        if sheet._cells.get((cell.row, cell.column)) is cell:
            edits[cell.coordinate] = _cell_edit(restorer, mark)
    return edits


def _cell_edit(restorer: "Restorer", mark: CellMark) -> CellEdit:
    cell = mark.cell
    unchanged = (cell.value, cell.data_type) == mark.value
    cm = str(restorer.dynamic_array) if mark.dynamic else mark.cm
    cm = cm if isinstance(cell.value, ArrayFormula) else None
    vm = mark.vm if unchanged else None
    cached = mark.cached if (unchanged or mark.dynamic) and cm else None

    def edit(attrs: str, body: str) -> tuple[str, str]:
        if cached:
            kind, text = cached
            attrs = _TYPE_ATTRIBUTE.sub("", attrs) + ("" if kind == "n" else f' t="{kind}"')
            body = _EMPTY_VALUE.sub("", body) + f"<v>{escape(text)}</v>"
        extra = (f' cm="{cm}"' if cm else "") + (f' vm="{vm}"' if vm else "")
        return attrs + extra, body

    return edit


def _drawing(restorer: "Restorer", sheet_part: str, package: SheetPackage) -> str | None:
    """Add the anchors openpyxl dropped to the sheet's drawing; the ``<drawing>`` element if new."""
    if not package.anchors:
        return None
    files = restorer.package
    rels = files.rels(sheet_part)
    rel = next((r for r in rels if r.type == REL_BASE + "drawing"), None)
    if rel is None:
        name = files.unique_name("xl/drawings/drawing1.xml")
        rel = Rel(free_id(rels, ""), REL_BASE + "drawing", name)
        rels.append(rel)
        files.content_types.declare(name, _DRAWING_TYPE)
        data, created = _EMPTY_DRAWING.encode(), f'<drawing xmlns:r="{REL_NS}" r:id="{rel.id}"/>'
    else:
        data, created = files.read(rel.target), None
    scope = package.drawing_namespaces
    anchors = restorer.attach(rel.target, package.drawing_links, package.anchors, scope)
    anchors = _renumber_shapes([with_namespaces(a, scope) for a in anchors], data)
    files.write(rel.target, append_to_root(data, "".join(anchors)))
    return created


def _vml(
    restorer: "Restorer", sheet_part: str, package: SheetPackage, document: Scan
) -> str | None:
    """Add the form controls and other shapes openpyxl dropped to the sheet's VML drawing.

    Returns the ``<legacyDrawing>`` element if the sheet had none, which is so without notes.
    """
    if not package.vml:
        return None
    files = restorer.package
    rels = files.rels(sheet_part)
    written = {
        i
        for c in document.children
        if c.local == "legacyDrawing"
        for i in relationship_ids(document.raw(c), {})
    }
    rel = next((r for r in rels if r.id in written), None)
    created = None
    if rel is None:
        name = files.unique_name("xl/drawings/vmlDrawing1.vml")
        rel = Rel(free_id(rels, ""), REL_BASE + "vmlDrawing", name)
        rels.append(rel)
        files.content_types.declare(name, _VML_TYPE, by_default=True)
        created = f'<legacyDrawing xmlns:r="{REL_NS}" r:id="{rel.id}"/>'
    existing = files.read(rel.target) if files.exists(rel.target) else None
    scope = package.vml_namespaces
    shapes = restorer.attach(rel.target, package.vml_links, package.vml, scope)
    files.write(rel.target, vml.merge(existing, shapes, scope))
    return created


def _pivot_extensions(restorer: "Restorer", sheet_part: str, package: SheetPackage) -> None:
    for rel in restorer.package.rels(sheet_part):
        if rel.type != REL_BASE + "pivotTable":
            continue
        name = restorer.pivot_name(rel.target)
        text = restorer.package.read(rel.target).decode("utf-8")
        for kind in pivots.FIELD_KINDS:
            found = {
                p: x for (n, k, p), x in package.pivot_fields.items() if (n, k) == (name, kind)
            }
            text = pivots.with_field_extensions(text, kind, found)
        data = text.encode("utf-8")
        if extension := package.pivot_extensions.get(name):
            data = append_to_root(data, extension)
        restorer.package.write(rel.target, data)


def _shifted(edits: list[Edit], offset: int) -> list[Edit]:
    return [(start - offset, end - offset, text) for start, end, text in edits]


def _renumber_shapes(anchors: list[str], drawing: bytes) -> list[str]:
    """Give preserved shapes new ids where the written drawing already uses theirs.

    Other parts name shapes by id (form controls, in the VML), so ids stay if they can.
    """
    used = {m[2] for m in _SHAPE_ID.finditer(drawing.decode("utf-8"))}
    following = max((int(i) for i in used), default=0)
    mapping: dict[str, str] = {}

    def renumber(match: re.Match[str]) -> str:
        nonlocal following
        if match[2] == "0" or match[2] not in used:
            return match[0]
        if match[2] not in mapping:
            following += 1
            mapping[match[2]] = str(following)
        return match[1] + mapping[match[2]] + match[3]

    return [_SHAPE_ID.sub(renumber, anchor) for anchor in anchors]
