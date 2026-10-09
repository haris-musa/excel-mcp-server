"""Excel 2016 charts (waterfall, histogram, ...) on a sheet: adding, listing, removing, copying.

openpyxl has no model for them. They live in the package layer as an anchor of the sheet's
drawing, a link to the chart part and the hidden names the part's data refers to.
"""

import re
import uuid
from dataclasses import dataclass
from html import escape, unescape
from pathlib import Path

from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from excel_mcp.operations import drawings
from excel_mcp.operations.chart_info import ChartInfo
from excel_mcp.operations.chartex_xml import KINDS, NameBook, chart_xml
from excel_mcp.operations.charts_data import Plot
from excel_mcp.operations.charts_options import ChartOptions
from excel_mcp.operations.sheet_refs import SheetCopyRefs
from excel_mcp.package import Link, Part, state_of
from excel_mcp.package.anchors import Geometry
from excel_mcp.package.model import SheetPackage
from excel_mcp.package.shape_names import anchored_name
from excel_mcp.refs import parse_cell

EMU_PER_CM = 360000
_URI = "http://schemas.microsoft.com/office/drawing/2014/chartex"
_CHART_EX = "application/vnd.ms-office.chartex+xml"
_CHART_REL = "http://schemas.microsoft.com/office/2014/relationships/chartEx"
_STYLE_REL = "http://schemas.microsoft.com/office/2011/relationships/chartStyle"
_COLORS_REL = "http://schemas.microsoft.com/office/2011/relationships/chartColorStyle"
_REQUIRED = {
    1: "http://schemas.microsoft.com/office/drawing/2015/9/8/chartex",
    2: "http://schemas.microsoft.com/office/drawing/2015/10/21/chartex",
}
_PARTS = Path(__file__).with_name("chartex_parts")
_CHART_ID = re.compile(r"<cx:chart\b[^>]*\br:id=\"([^\"]+)\"")
_FORMULA = re.compile(r"(<cx:f\b[^>]*>)([^<]*)(</cx:f>)")
_NAME = re.compile(r"_xlchart\.v(\d+)\.\d+")
_TITLE = re.compile(r"<cx:chart>\s*<cx:title\b.*?<cx:v>(.*?)</cx:v>.*?<cx:plotArea", re.S)
_DATA = re.compile(r'<cx:data id="(\d+)">.*?<cx:numDim\b[^>]*><cx:f\b[^>]*>(.*?)</cx:f>', re.S)
_LAYOUT = re.compile(r'<cx:series layoutId="(\w+)"')
_TYPES = {"clusteredColumn": "histogram", "boxWhisker": "box_whisker"}
_FALLBACK = (
    "This chart isn't available in your version of Excel.\n\nEditing this shape or saving "
    "this workbook into a different file format will permanently break the chart."
)


@dataclass(frozen=True)
class Modern:
    """A chart of the sheet: its anchor, and the link to its chart part."""

    anchor: str
    link: Link

    @property
    def part(self) -> Part:
        assert isinstance(self.link.target, Part)
        return self.link.target

    @property
    def xml(self) -> str:
        return self.part.data.decode("utf-8")

    @property
    def name(self) -> str:
        return anchored_name(self.anchor) or ""


def _workbook(sheet: Worksheet) -> Workbook:
    assert sheet.parent is not None
    return sheet.parent


def modern_charts(sheet: Worksheet) -> list[Modern]:
    package = state_of(_workbook(sheet)).sheet(sheet)
    links = {link.id: link for link in package.drawing_links}
    found = []
    for anchor in package.anchors:
        match = _CHART_ID.search(anchor) if _URI in anchor else None
        if match and match[1] in links and isinstance(links[match[1]].target, Part):
            found.append(Modern(anchor, links[match[1]]))
    return found


def describe(sheet: Worksheet) -> list[ChartInfo]:
    names = _workbook(sheet).defined_names
    infos = []
    for chart in modern_charts(sheet):
        xml = chart.xml
        title = _TITLE.search(xml)
        series = [names[n].attr_text if n in names else n for _, n in _DATA.findall(xml)]
        infos.append(
            ChartInfo(
                name=chart.name,
                type=_type(xml),
                title=unescape(title[1]) if title else None,
                range=drawings.anchored_range(chart.anchor),
                series=[s or "" for s in series],
            )
        )
    return infos


def _type(xml: str) -> str:
    layout = _LAYOUT.search(xml)
    name = layout[1] if layout else ""
    if "paretoLine" in xml:
        return "pareto"
    return _TYPES.get(name, name)


def add(
    workbook: Workbook,
    sheet: Worksheet,
    anchor_cell: str,
    chart_type: str,
    plots: list[Plot],
    options: ChartOptions,
    replacing: Modern | None,
    name: str,
) -> str:
    """Put a chart on the sheet: in the place of `replacing` if given, else after the others.
    Returns the cells it covers."""
    package = state_of(workbook).sheet(sheet)
    position = len(package.anchors)
    if replacing is not None:
        position = package.anchors.index(replacing.anchor)
        remove(sheet, replacing)
    number = len(sheet._charts) + len(modern_charts(sheet)) + 1  # pyright: ignore[reportAttributeAccessIssue]
    link = Link(_free_id(package), _CHART_REL, _chart_part(workbook, chart_type, plots, options))
    package.drawing_links.append(link)
    version = KINDS[chart_type].version
    anchor = _anchor(sheet, anchor_cell, options, number, link.id, version, name)
    package.anchors.insert(position, anchor)
    return drawings.anchored_range(anchor) or ""


def remove(sheet: Worksheet, chart: Modern) -> None:
    package = state_of(_workbook(sheet)).sheet(sheet)
    package.anchors.remove(chart.anchor)
    package.drawing_links.remove(chart.link)
    for _, name, _ in _FORMULA.findall(chart.xml):
        if _NAME.fullmatch(name):
            _workbook(sheet).defined_names.pop(name, None)


def copy(source: Worksheet, target: Worksheet, refs: SheetCopyRefs) -> None:
    """Copy the charts of `source` to `target`, with their data pointing at the copy."""
    workbook = _workbook(source)
    package = state_of(workbook).sheet(target)
    book = NameBook(workbook)
    for chart in modern_charts(source):

        def repoint(match: re.Match[str]) -> str:
            text = match[2]
            if found := _NAME.fullmatch(text):
                reference = workbook.defined_names[text].attr_text or ""
                text = book.of(refs.operand(reference, target.title), int(found[1]))
            else:
                text = refs.operand(text, target.title)
            return f"{match[1]}{text}{match[3]}"

        data = _FORMULA.sub(repoint, chart.xml).encode("utf-8")
        part = Part(chart.part.name, _CHART_EX, data, links=_cloned(chart.part.links))
        link = Link(chart.link.id, chart.link.type, part)
        package.drawing_links.append(link)
        package.anchors.append(chart.anchor)


def _cloned(links: list[Link]) -> list[Link]:
    return [
        Link(link.id, link.type, Part(t.name, t.content_type, t.data, by_default=t.by_default))
        for link in links
        if isinstance(t := link.target, Part)
    ]


def _free_id(package: SheetPackage) -> str:
    used = [int(link.id[3:]) for link in package.drawing_links if link.id[3:].isdigit()]
    return f"rId{max(used, default=0) + 1}"


def _chart_part(
    workbook: Workbook, chart_type: str, plots: list[Plot], options: ChartOptions
) -> Part:
    style = _file(f"{'histogram' if chart_type == 'pareto' else chart_type}_style.xml")
    return Part(
        "xl/charts/chartEx1.xml",
        _CHART_EX,
        chart_xml(workbook, chart_type, plots, options).encode("utf-8"),
        links=[
            Link("rId1", _STYLE_REL, _style(style)),
            Link("rId2", _COLORS_REL, _colors()),
        ],
    )


def _file(name: str) -> bytes:
    return (_PARTS / name).read_bytes()


def _style(data: bytes) -> Part:
    return Part("xl/charts/style1.xml", "application/vnd.ms-office.chartstyle+xml", data)


def _colors() -> Part:
    return Part(
        "xl/charts/colors1.xml",
        "application/vnd.ms-office.chartcolorstyle+xml",
        _file("colors.xml"),
    )


def _anchor(
    sheet: Worksheet,
    anchor_cell: str,
    options: ChartOptions,
    number: int,
    rid: str,
    version: int,
    name: str,
) -> str:
    row, col = parse_cell(anchor_cell)
    geometry = Geometry(sheet)
    width, height = round(options.width_cm * EMU_PER_CM), round(options.height_cm * EMU_PER_CM)
    last_col, col_off = _walk(geometry, False, col, width)
    last_row, row_off = _walk(geometry, True, row, height)
    x = sum(geometry.size(False, n) for n in range(1, col))
    y = sum(geometry.size(True, n) for n in range(1, row))
    return (
        '<xdr:twoCellAnchor xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/'
        'spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        f"<xdr:from><xdr:col>{col - 1}</xdr:col><xdr:colOff>0</xdr:colOff>"
        f"<xdr:row>{row - 1}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>"
        f"<xdr:to><xdr:col>{last_col}</xdr:col><xdr:colOff>{col_off}</xdr:colOff>"
        f"<xdr:row>{last_row}</xdr:row><xdr:rowOff>{row_off}</xdr:rowOff></xdr:to>"
        '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">'
        f'<mc:Choice xmlns:cx{version}="{_REQUIRED[version]}" Requires="cx{version}">'
        '<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr>'
        f'<xdr:cNvPr id="{number + 1}" name="{escape(name)}"><a:extLst>'
        '<a:ext uri="{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}"><a16:creationId '
        'xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main" '
        f'id="{{{str(uuid.uuid4()).upper()}}}"/></a:ext></a:extLst></xdr:cNvPr>'
        "<xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>"
        '<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>'
        f'<a:graphic><a:graphicData uri="{_URI}"><cx:chart '
        'xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        f'r:id="{rid}"/></a:graphicData></a:graphic></xdr:graphicFrame></mc:Choice>'
        '<mc:Fallback><xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="0" name=""/>'
        '<xdr:cNvSpPr><a:spLocks noTextEdit="1"/></xdr:cNvSpPr></xdr:nvSpPr><xdr:spPr>'
        f'<a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{width}" cy="{height}"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:prstClr val="white"/>'
        '</a:solidFill><a:ln w="1"><a:solidFill><a:prstClr val="green"/></a:solidFill></a:ln>'
        '</xdr:spPr><xdr:txBody><a:bodyPr vertOverflow="clip" horzOverflow="clip"/><a:lstStyle/>'
        f'<a:p><a:r><a:rPr lang="en-US" sz="1100"/><a:t>{_FALLBACK}</a:t></a:r></a:p>'
        "</xdr:txBody></xdr:sp></mc:Fallback></mc:AlternateContent><xdr:clientData/>"
        "</xdr:twoCellAnchor>"
    )


def _walk(geometry: Geometry, rows: bool, line: int, distance: int) -> tuple[int, int]:
    """The 0-based line and offset where `distance` EMU from the top of `line` ends."""
    while distance >= geometry.size(rows, line):
        distance -= geometry.size(rows, line)
        line += 1
    return line - 1, distance
