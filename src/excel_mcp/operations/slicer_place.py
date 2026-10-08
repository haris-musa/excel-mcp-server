"""Putting slicers and timelines into a workbook's package: parts, lists, shapes and names."""

import re
import uuid

from openpyxl.workbook import Workbook
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.worksheet import Worksheet

from excel_mcp.operations import slicer_xml as xml
from excel_mcp.package import Link, Part, SheetPackage, state_of
from excel_mcp.package.anchors import Geometry
from excel_mcp.package.consistency import SLICER_CACHE, TIMELINE_CACHE
from excel_mcp.package.scan import relationship_ids

_XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
_A14 = "http://schemas.microsoft.com/office/drawing/2010/main"
_SLE = "http://schemas.microsoft.com/office/drawing/2010/slicer"
_SLE15 = "http://schemas.microsoft.com/office/drawing/2012/slicer"
_TSLE = "http://schemas.microsoft.com/office/drawing/2012/timeslicer"
_A16 = "http://schemas.microsoft.com/office/drawing/2014/main"
_SHAPE_ID = re.compile(r'<(?:\w+:)?cNvPr\b[^>]*?\bid="(\d+)"')
_FALLBACK = {
    "pivot": "This shape represents a slicer. Slicers are supported in Excel 2010 or later.\n\n"
    "If the shape was modified in an earlier version of Excel, or if the workbook was saved in "
    "Excel 2003 or earlier, the slicer cannot be used.",
    "table": "This shape represents a table slicer. Table slicers are not supported in this "
    "version of Excel.\n\nIf the shape was modified in an earlier version of Excel, or if the "
    "workbook was saved in Excel 2007 or earlier, the slicer can't be used.",
    "timeline": "Timeline: Works in Excel 2013 or higher. Do not move or resize.",
}


def add_cache(workbook: Workbook, kind: str, name: str, text: str) -> None:
    """Add a slicer or timeline cache, with its entry in the workbook's extension list."""
    package = state_of(workbook).workbook
    timeline = kind == "timeline"
    folder = "timeline" if timeline else "slicer"
    part = Part(
        f"xl/{folder}Caches/{folder}Cache1.xml",
        TIMELINE_CACHE if timeline else SLICER_CACHE,
        text.encode("utf-8"),
    )
    rid = _free_id(package.links, "rIdCache")
    package.links.append(
        Link(rid, xml.TIMELINE_CACHE_REL if timeline else xml.SLICER_CACHE_REL, part)
    )
    uri, element = xml.workbook_cache_extensions(kind, rid)
    _list_in(package.extensions, uri, element)
    package.extensions.setdefault(xml.WORKBOOK_PR, xml.WORKBOOK_PR_EXTENSION)
    workbook.defined_names[name] = DefinedName(name, attr_text="#N/A")


def add_entry(workbook: Workbook, sheet: Worksheet, kind: str, entry: str, shape: str) -> None:
    """Add a slicer or timeline to a sheet's list of them, and its shape to the drawing."""
    package = state_of(workbook).sheet(sheet)
    timeline = kind == "timeline"
    uri, element = xml.sheet_slicer_extension(kind, "")
    existing = _part_listed(package, uri)
    if existing is not None:
        text = existing.data.decode("utf-8")
        closing = "</timelines>" if timeline else "</slicers>"
        existing.data = text.replace(closing, entry + closing).encode("utf-8")
    else:
        wrap = xml.timelines_part if timeline else xml.slicers_part
        name = "timelines/timeline1.xml" if timeline else "slicers/slicer1.xml"
        part = Part(
            f"xl/{name}",
            xml.TIMELINE_PART if timeline else xml.SLICER_PART,
            wrap(entry).encode("utf-8"),
        )
        rid = _free_id(package.links, "rIdSlicers")
        package.links.append(Link(rid, xml.relationship_type(kind), part))
        _, element = xml.sheet_slicer_extension(kind, rid)
        package.set_extension(uri, element)
    package.anchors.append(shape)


def _part_listed(package: SheetPackage, uri: str) -> Part | None:
    listed = package.extensions.get(uri)
    if listed is None:
        return None
    wanted = relationship_ids(listed, package.namespaces)
    link = next((link for link in package.links if link.id in wanted), None)
    return link.target if link and isinstance(link.target, Part) else None


def _list_in(extensions: dict[str, str], uri: str, element: str) -> None:
    """Add the reference inside ``element`` to the list under ``uri``, or the whole entry."""
    if uri not in extensions:
        extensions[uri] = element
        return
    reference = re.search(r"<(?:\w+:)(?:slicerCache|timelineCacheRef)\b[^>]*/>", element)
    assert reference is not None
    existing = extensions[uri]
    closing = re.search(r"</(?:\w+:)(?:slicerCaches|timelineCacheRefs)>", existing)
    assert closing is not None
    extensions[uri] = existing[: closing.start()] + reference[0] + existing[closing.start() :]


def _free_id(links: list[Link], stem: str) -> str:
    taken = {link.id for link in links}
    number = 1
    while f"{stem}{number}" in taken:
        number += 1
    return f"{stem}{number}"


def next_shape_id(package: SheetPackage) -> int:
    used = [int(i) for anchor in package.anchors for i in _SHAPE_ID.findall(anchor)]
    return max(used, default=1) + 1


def anchor(
    sheet: Worksheet,
    kind: str,
    name: str,
    first: tuple[int, int],
    size: tuple[int, int],
    shape_id: int,
) -> str:
    """The shape of a slicer or timeline: a two-cell anchor with a fallback for old Excel."""
    geometry = Geometry(sheet)
    row, column = first
    left = sum(geometry.size(False, c) for c in range(1, column))
    top = sum(geometry.size(True, r) for r in range(1, row))
    end_col, col_off = _end(geometry, False, column, left + size[0], left)
    end_row, row_off = _end(geometry, True, row, top + size[1], top)
    timeline = kind == "timeline"
    requires = "tsle" if timeline else "sle15" if kind == "table" else "a14"
    declared = {"tsle": _TSLE, "sle15": _SLE15, "a14": _A14}[requires]
    graphic = (
        f'<tsle:timeslicer xmlns:tsle="{_TSLE}" name="{name}"/>'
        if timeline
        else f'<sle:slicer xmlns:sle="{_SLE}" name="{name}"/>'
    )
    uri = _TSLE if timeline else _SLE
    guid = "{" + str(uuid.uuid4()).upper() + "}"
    edit = "absolute" if kind == "table" else "oneCell"
    return (
        f'<xdr:twoCellAnchor xmlns:xdr="{_XDR}" xmlns:a="{_A}" editAs="{edit}">'
        f"<xdr:from><xdr:col>{column - 1}</xdr:col><xdr:colOff>0</xdr:colOff>"
        f"<xdr:row>{row - 1}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>"
        f"<xdr:to><xdr:col>{end_col}</xdr:col><xdr:colOff>{col_off}</xdr:colOff>"
        f"<xdr:row>{end_row}</xdr:row><xdr:rowOff>{row_off}</xdr:rowOff></xdr:to>"
        f'<mc:AlternateContent xmlns:mc="{xml.MC}">'
        f'<mc:Choice xmlns:{requires}="{declared}" Requires="{requires}">'
        f'<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr>'
        f'<xdr:cNvPr id="{shape_id}" name="{name}"><a:extLst>'
        f'<a:ext uri="{{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}}">'
        f'<a16:creationId xmlns:a16="{_A16}" id="{guid}"/></a:ext></a:extLst></xdr:cNvPr>'
        f"<xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>"
        f'<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>'
        f'<a:graphic><a:graphicData uri="{uri}">{graphic}</a:graphicData></a:graphic>'
        f"</xdr:graphicFrame></mc:Choice>"
        f'<mc:Fallback><xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="0" name=""/>'
        f'<xdr:cNvSpPr><a:spLocks noTextEdit="1"/></xdr:cNvSpPr></xdr:nvSpPr><xdr:spPr>'
        f'<a:xfrm><a:off x="{left}" y="{top}"/><a:ext cx="{size[0]}" cy="{size[1]}"/></a:xfrm>'
        f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        f'<a:solidFill><a:prstClr val="white"/></a:solidFill>'
        f'<a:ln w="1"><a:solidFill><a:prstClr val="green"/></a:solidFill></a:ln></xdr:spPr>'
        f'<xdr:txBody><a:bodyPr vertOverflow="clip" horzOverflow="clip"/><a:lstStyle/>'
        f'<a:p><a:r><a:rPr lang="en-US" sz="1100"/><a:t>{_FALLBACK[kind]}</a:t></a:r></a:p>'
        f"</xdr:txBody></xdr:sp></mc:Fallback></mc:AlternateContent><xdr:clientData/>"
        f"</xdr:twoCellAnchor>"
    )


def _end(geometry: Geometry, rows: bool, start: int, end: int, origin: int) -> tuple[int, int]:
    """The line (0-based) and offset where ``end`` EMU from the sheet's edge falls."""
    line, position = start, origin
    while position + geometry.size(rows, line) <= end:
        position += geometry.size(rows, line)
        line += 1
    return line - 1, end - position
