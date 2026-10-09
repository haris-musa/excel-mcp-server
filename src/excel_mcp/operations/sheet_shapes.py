"""Copying what openpyxl has no model for in the drawing of a sheet: shapes, text boxes,
connectors, groups and form controls."""

import re
import uuid
from collections.abc import Callable
from xml.sax.saxutils import escape, quoteattr

from openpyxl.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from excel_mcp.operations.sheet_refs import SheetCopyRefs
from excel_mcp.operations.slicer_xml import PIVOT_SLICERS, TABLE_SLICERS, TIMELINES
from excel_mcp.package import PackageState, Part, SheetPackage, state_of
from excel_mcp.package.consistency import THREADED_COMMENTS
from excel_mcp.package.extensions import CONDITIONAL_FORMATS, SPARKLINES
from excel_mcp.package.model import cloned
from excel_mcp.package.scan import relationship_ids, unescape

# Charts of Excel 2016 and slicers are copied with the parts they need, by their own code.
_COPIED_ELSEWHERE = (
    "drawing/2014/chartex",
    "drawing/2010/slicer",
    "drawing/2012/timeslicer",
)
_CREATION_ID = re.compile(r'(<a16:creationId\b[^>]*?\bid=")[^"]*(")')
_CONTROL_ID = re.compile(r'\bshapeId="(\d+)"')
_SHAPE_NUMBER = re.compile(r"_x0000_s(\d+)")
_VML_SHAPE = re.compile(r"<(?:[\w.-]+:)?shape\b")
_CONTROL_ANCHOR = re.compile(r'compatExt\b[^>]*\bspid="_x0000_s(\d+)"')
_PROPERTY_FORMULA = re.compile(r'(\bfmla(?:Link|Range|Group|Txbx)=)"([^"]*)"')
_VML_FORMULA = re.compile(r"(<x:Fmla(?:Link|Range|Group|Txbx)>)([^<]*)(</x:Fmla\w+>)")
_OBJECTS = ("controls", "oleObjects")  # objects with a shape in the VML drawing
_CONTROL_PROPERTIES = "application/vnd.ms-excel.controlproperties+xml"
_IDS_PER_SHEET = 1024
# What holds a shape's id, and how many ids make one step of the number there.
_NUMBERS = (
    (re.compile(r"(_x0000_s)(\d+)()"), 1),
    (re.compile(r'(\bshapeId=")(\d+)(")'), 1),
    (re.compile(r'(<(?:[\w.-]+:)?cNvPr\b[^>]*?\bid=")(\d+)(")'), 1),
    (re.compile(r'(<(?:[\w.-]+:)?idmap\b[^>]*?\bdata=")(\d+)(")'), _IDS_PER_SHEET),
)


def copy_shapes(
    workbook: Workbook, source: Worksheet, target: Worksheet, refs: SheetCopyRefs
) -> None:
    """Give the copy of a sheet the shapes and form controls of the original, with the
    pictures they use. What a control reads or sets on the original it does on the copy."""
    state = state_of(workbook)
    package = state.sheets.get(source)
    if package is None:
        return
    copy = state.sheet(target)
    controls = _control_ids(package)
    shift = _own_numbers(state, controls)
    used: set[str] = set()
    for anchor in package.anchors:
        if any(uri in anchor for uri in _COPIED_ELSEWHERE):
            continue
        found = _CONTROL_ANCHOR.search(anchor)
        if found and found[1] not in controls:
            continue
        anchor = _shifted(anchor, shift) if found else anchor
        copy.anchors.append(_CREATION_ID.sub(_new_id, anchor))
        used |= relationship_ids(anchor, package.drawing_namespaces)
    copy.drawing_links += cloned([link for link in package.drawing_links if link.id in used])
    if controls:
        repoint = lambda text: refs.operand(text, target.title)  # noqa: E731
        _copy_controls(package, copy, controls, shift, repoint)


def _control_ids(package: SheetPackage) -> set[str]:
    return {
        i for name, xml in package.elements if name in _OBJECTS for i in _CONTROL_ID.findall(xml)
    }


def _own_numbers(state: PackageState, controls: set[str]) -> int:
    """How far to move the ids of the controls of a copy: Excel gives each sheet its own block
    of 1024 ids for them (1025 to 2048 on the first, 2049 to 3072 on the second), and the sheet
    shows the controls of other sheets that have the same ids."""
    if not controls:
        return 0
    block = {
        (int(i) - 1) // _IDS_PER_SHEET
        for package in state.sheets.values()
        for i in _control_ids(package)
    }
    return (max(block) + 1 - (min(int(i) for i in controls) - 1) // _IDS_PER_SHEET) * _IDS_PER_SHEET


def _shifted(xml: str, shift: int) -> str:
    for pattern, step in _NUMBERS:
        xml = pattern.sub(lambda m, step=step: f"{m[1]}{int(m[2]) + shift // step}{m[3]}", xml)
    return xml


def _copy_controls(
    package: SheetPackage,
    copy: SheetPackage,
    controls: set[str],
    shift: int,
    repoint: Callable[[str], str],
) -> None:
    """The ``<controls>`` element, the property parts it names, and the controls' VML shapes."""
    elements = [(name, xml) for name, xml in package.elements if name in _OBJECTS]
    copy.elements += [(name, _shifted(xml, shift)) for name, xml in elements]
    named = {i for _, xml in elements for i in relationship_ids(xml, package.namespaces)}
    for link in cloned([link for link in package.links if link.id in named]):
        if isinstance(link.target, Part):
            link.target.unique = True
            if link.target.content_type == _CONTROL_PROPERTIES:
                link.target.data = _PROPERTY_FORMULA.sub(
                    lambda m: f"{m[1]}{quoteattr(repoint(unescape(m[2])))}",
                    link.target.data.decode("utf-8"),
                ).encode("utf-8")
        copy.links.append(link)
    shapes = [
        s
        for s in package.vml
        if not _VML_SHAPE.match(s) or set(_SHAPE_NUMBER.findall(s)) & controls
    ]
    copy.vml += [
        _shifted(
            _VML_FORMULA.sub(lambda m: m[1] + escape(repoint(unescape(m[2]))) + m[3], shape), shift
        )
        for shape in shapes
    ]
    wanted = {i for shape in shapes for i in relationship_ids(shape, package.vml_namespaces)}
    copy.vml_links += cloned([link for link in package.vml_links if link.id in wanted])


def _new_id(match: re.Match[str]) -> str:
    return f"{match[1]}{{{str(uuid.uuid4()).upper()}}}{match[2]}"


_HANDLED_EXTENSIONS = {SPARKLINES, CONDITIONAL_FORMATS, PIVOT_SLICERS, TABLE_SLICERS, TIMELINES}


def copy_sheet_elements(workbook: Workbook, source: Worksheet, target: Worksheet) -> list[str]:
    """Copy the other sheet-level elements openpyxl drops (protected ranges, ignored errors,
    sort state, ...). Returns what cannot be copied, for the tool to report."""
    package = state_of(workbook).sheets.get(source)
    if package is None:
        return []
    copy = state_of(workbook).sheet(target)
    skipped = []
    for name, xml in package.elements:
        if name in _OBJECTS or name == "ignoredErrors":  # Excel's copy has none of the latter
            continue
        copy.elements.append((name, xml))
        named = relationship_ids(xml, package.namespaces)
        for link in cloned([link for link in package.links if link.id in named]):
            if isinstance(link.target, Part):
                link.target.unique = True
            copy.links.append(link)
    others = len(set(package.extensions) - _HANDLED_EXTENSIONS)
    if others:
        skipped.append(f"{others} other sheet extension{'s' if others != 1 else ''}")
    if any(
        isinstance(link.target, Part) and link.target.content_type == THREADED_COMMENTS
        for link in package.links
    ):
        skipped.append("threaded comments (the notes are copied)")
    return skipped
