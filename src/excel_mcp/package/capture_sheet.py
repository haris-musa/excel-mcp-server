"""Reading from the original file what openpyxl is going to drop from a worksheet."""

import re
from typing import Literal
from xml.etree import ElementTree

from openpyxl.cell.cell import Cell
from openpyxl.utils.cell import coordinate_to_tuple
from openpyxl.worksheet.formula import ArrayFormula
from openpyxl.worksheet.worksheet import Worksheet

from excel_mcp.package import vml
from excel_mcp.package.model import CellMark, SheetPackage
from excel_mcp.package.opc import REL_BASE, Rel
from excel_mcp.package.pivots import FIELD_KINDS
from excel_mcp.package.pivots import fields as _fields_in
from excel_mcp.package.reader import Reader
from excel_mcp.package.scan import (
    REL_NS,
    attributes,
    relationship_ids,
    scan,
    scan_fragment,
    unescape,
    with_namespaces,
)
from excel_mcp.package.shape_names import anchored_name, restore_names
from excel_mcp.refs import CellRange, parse_range

MARKUP_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"
_XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
_DRAWINGML = "http://schemas.openxmlformats.org/drawingml/2006/main"
_CHART_DATA = "http://schemas.openxmlformats.org/drawingml/2006/chart"

# Elements openpyxl never writes. Ones it models are left to it, even when it omits them
# on purpose (a tool removed the sheet protection, say).
SHEET_ELEMENTS = frozenset(
    [
        "sheetCalcPr",
        "protectedRanges",
        "sortState",
        "dataConsolidate",
        "customSheetViews",
        "phoneticPr",
        "customProperties",
        "cellWatches",
        "ignoredErrors",
        "smartTags",
        "legacyDrawingHF",
        "drawingHF",
        "picture",
        "oleObjects",
        "controls",
        "webPublishItems",
    ]
)
# Relationships openpyxl reads and writes again itself (or deliberately drops).
_SHEET_RELS = {
    REL_BASE + name
    for name in ["hyperlink", "drawing", "comments", "table", "pivotTable", "printerSettings"]
}
_MARKED_CELL = re.compile(rb'<c\s[^>]*?\s(?:cm|vm)="[^"]*"[^>]*?(/?)>')
_VALUE = re.compile(r"<v>(.*?)</v>", re.S)
_EXTENSION_LIST = re.compile(r"<extLst>.*</extLst>", re.S)


def capture_sheet(reader: Reader, name: str, sheet: Worksheet, keep_vml: bool) -> SheetPackage:
    """What openpyxl drops of a sheet. In a macro-enabled workbook it keeps the VML itself."""
    data = reader.archive.read(name)
    document = scan(data, "sheetData")
    package = SheetPackage(namespaces=document.namespaces)
    legacy = ""
    for child in document.children:
        raw = document.raw(child)
        if child.local == "extLst":
            package.extensions = entries(raw, document.namespaces)
        elif child.local == "conditionalFormatting":
            package.rule_extensions.update(_rule_extensions(raw, document.namespaces))
        elif child.local == "legacyDrawing":
            legacy = raw
        elif child.local in SHEET_ELEMENTS or child.uri == MARKUP_NS:
            local = "controls" if child.uri == MARKUP_NS else child.local
            package.elements.append((local, with_namespaces(raw, document.namespaces)))
    referenced = set().union(
        *(relationship_ids(xml, document.namespaces) for _, xml in package.elements),
        *(relationship_ids(xml, document.namespaces) for xml in package.extensions.values()),
    )
    comments = relationship_ids(legacy, document.namespaces)
    package.links = reader.links(name, lambda rel: _keeps(rel, referenced, comments))
    if b' cm="' in data or b' vm="' in data:
        package.marks = _marks(data, sheet)
    for rel in reader.rels(name):
        if rel.id in comments and rel.target in reader.names:
            if keep_vml:
                _vml(reader, rel.target, package)
        elif rel.type == REL_BASE + "drawing" and rel.target in reader.names:
            _drawing(reader, rel.target, package, sheet)
        elif rel.type == REL_BASE + "pivotTable" and rel.target in reader.names:
            _pivot(reader.archive.read(rel.target), package)
    return package


def _keeps(rel: Rel, referenced: set[str], comments: set[str]) -> bool:
    if rel.id in referenced:
        return True
    return rel.id not in comments and rel.type not in _SHEET_RELS


def _rule_extensions(block: str, namespaces: dict[str, str]) -> dict[tuple[str, str], str]:
    """The ``<extLst>`` of each rule of a conditional format by (range, priority)."""
    sqref = attributes(block[: block.index(">") + 1]).get("sqref", "")
    found = {}
    rules = scan_fragment(block, namespaces)
    for rule in rules.children:
        text = rules.raw(rule)
        priority = attributes(text[: text.index(">") + 1]).get("priority", "")
        if extension := part_extension(with_namespaces(text, namespaces).encode()):
            found[(sqref, priority)] = extension
    return found


def _marks(data: bytes, sheet: Worksheet) -> list[CellMark]:
    marks = []
    for found in _MARKED_CELL.finditer(data):
        values = attributes(found.group(0).decode("utf-8"))
        cell = sheet._cells.get(coordinate_to_tuple(values["r"]))
        if not isinstance(cell, Cell):
            continue
        mark = CellMark(
            cell, values.get("cm"), values.get("vm"), value=(cell.value, cell.data_type)
        )
        if mark.cm and not found.group(1):
            body = data[found.end() : data.index(b"</c>", found.end())].decode("utf-8")
            result = _VALUE.search(body)
            if "<f" in body and result:
                mark.cached = (values.get("t", "n"), unescape(result[1]))
        if mark.cm and isinstance(cell.value, ArrayFormula):
            mark.spill = _spilled(sheet, cell, parse_range(cell.value.ref))
        marks.append(mark)
    return marks


def _spilled(sheet: Worksheet, anchor: Cell, area: CellRange) -> list[tuple[Cell, object]]:
    """The cells of an array formula's range that hold the values it spilled."""
    spilled = []
    for row in range(area.min_row, area.max_row + 1):
        for col in range(area.min_col, area.max_col + 1):
            cell = sheet._cells.get((row, col))
            if isinstance(cell, Cell) and cell is not anchor and cell.value is not None:
                spilled.append((cell, (cell.value, cell.data_type)))
    return spilled


def _drawing(reader: Reader, name: str, package: SheetPackage, sheet: Worksheet) -> None:
    document = scan(reader.archive.read(name))
    package.drawing_namespaces = document.namespaces
    needed: set[str] = set()
    names: dict[str, dict[str, list[str | None]]] = {"chart": {}, "picture": {}}
    for child in document.children:
        raw = document.raw(child)
        kind = _modeled(raw, document.namespaces)
        if kind is None:
            package.anchors.append(raw)
            needed |= relationship_ids(raw, document.namespaces)
        else:
            names[kind].setdefault(child.local, []).append(anchored_name(raw))
    package.drawing_links = reader.links(name, lambda rel: rel.id in needed)
    restore_names(sheet._charts, _in_openpyxl_order(names["chart"]))  # pyright: ignore[reportAttributeAccessIssue]
    restore_names(sheet._images, _in_openpyxl_order(names["picture"]))  # pyright: ignore[reportAttributeAccessIssue]


def _in_openpyxl_order(by_anchor: dict[str, list[str | None]]) -> list[str | None]:
    """openpyxl lists absolute anchors first, then one-cell, then two-cell ones."""
    return [
        n
        for tag in ("absoluteAnchor", "oneCellAnchor", "twoCellAnchor")
        for n in by_anchor.get(tag, [])
    ]


def _vml(reader: Reader, name: str, package: SheetPackage) -> None:
    document = scan(reader.archive.read(name))
    package.vml = vml.preserved_shapes(document)
    package.vml_namespaces = document.namespaces
    needed = {i for xml in package.vml for i in relationship_ids(xml, document.namespaces)}
    package.vml_links = reader.links(name, lambda rel: rel.id in needed)


def _modeled(anchor: str, namespaces: dict[str, str]) -> Literal["chart", "picture"] | None:
    """What openpyxl turns an anchor into, and so writes back; None if it drops it."""
    declared = " ".join(f'xmlns{":" + p if p else ""}="{u}"' for p, u in namespaces.items())
    element = ElementTree.fromstring(f"<w {declared}>{anchor}</w>")[0]
    for child in element:
        if child.tag == f"{{{_XDR}}}graphicFrame":
            data = child.find(f".//{{{_DRAWINGML}}}graphicData")
            return "chart" if data is not None and data.get("uri") == _CHART_DATA else None
        if child.tag == f"{{{_XDR}}}pic":
            blip = child.find(f".//{{{_DRAWINGML}}}blip")
            return "picture" if blip is not None and blip.get(f"{{{REL_NS}}}embed") else None
    return None


def _pivot(data: bytes, package: SheetPackage) -> None:
    document = scan(data)
    name = attributes(data[document.root_tag[0] : document.root_tag[1]].decode("utf-8"))["name"]
    if extension := part_extension(data):
        package.pivot_extensions[name] = extension
    for kind in FIELD_KINDS:
        for position, field in enumerate(_fields_in(data.decode("utf-8"), kind)):
            if extension := _EXTENSION_LIST.search(field):
                text = with_namespaces(extension[0], document.namespaces)
                package.pivot_fields[name, kind, position] = text


def entries(extension_list: str, namespaces: dict[str, str]) -> dict[str, str]:
    """The entries of an ``<extLst>`` by uri."""
    document = scan_fragment(extension_list, namespaces)
    found = {}
    for entry in document.children:
        raw = document.raw(entry)
        found[attributes(raw[: raw.index(">") + 1]).get("uri", "")] = with_namespaces(
            raw, namespaces
        )
    return found


def part_extension(data: bytes) -> str | None:
    """The ``<extLst>`` that closes a part, which openpyxl writes without its content."""
    document = scan(data)
    for child in document.children:
        if child.local == "extLst":
            return with_namespaces(document.raw(child), document.namespaces)
    return None
