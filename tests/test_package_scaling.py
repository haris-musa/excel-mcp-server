"""Preserved content is processed in one pass over its XML, whatever the number of items.

A workbook can hold tens of thousands of rules, relationships or cells that the package layer
has to carry. Work that rescans a whole document per item takes quadratic time, so these tests
give inputs far too large for that and a generous time limit, with the result checked as well.
"""

import io
import re
import time
import zipfile
from collections.abc import Callable
from pathlib import Path

import pytest
from openpyxl import Workbook

from excel_mcp.package import consistency
from excel_mcp.package import extensions as ext
from excel_mcp.package.lines import LineEdit
from excel_mcp.package.model import Link, Part
from excel_mcp.package.patch import patch_cells
from excel_mcp.refs import parse_range
from tests.conftest import ToolCall
from tests.package_support import FIXTURES, read_parts, text

pytestmark = pytest.mark.anyio

LIMIT = 3.0
X14 = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
XM = "http://schemas.microsoft.com/office/excel/2006/main"
RULE_URI = "{B025F937-C7B1-47D3-B67F-A62EFF666E3E}"
FORMATS_URI = "{78C0D931-6437-407d-A8EE-F0AAD7539E65}"


def _previous_without_rules(extensions: dict[str, str], identifiers: set[str]) -> None:
    """`without_rules` as it was when it rescanned the XML once per identifier."""
    xml = extensions.get(ext.CONDITIONAL_FORMATS)
    if xml is None:
        return
    for identifier in sorted(identifiers):
        pattern = (
            rf'<x14:cfRule\b(?=[^>]*\bid="{re.escape(identifier)}")'
            r"(?:[^>]*/>|[^>]*>.*?</x14:cfRule>)"
        )
        xml = re.sub(pattern, "", xml, flags=re.S)
    emptied = (
        r"<x14:conditionalFormatting\b[^>]*>(?:(?!<x14:cfRule\b).)*?"
        r"</x14:conditionalFormatting>"
    )
    xml = re.sub(emptied, "", xml, flags=re.S)
    if "<x14:cfRule" in xml:
        extensions[ext.CONDITIONAL_FORMATS] = xml
    else:
        del extensions[ext.CONDITIONAL_FORMATS]


def _excel_extensions() -> dict[str, str]:
    """The extensions of a sheet as Excel 2010 wrote them: a data bar and an icon set."""
    with zipfile.ZipFile(FIXTURES / "excel_sparklines.xlsx") as archive:
        xml = archive.read("xl/worksheets/sheet2.xml").decode("utf-8")
    return {
        m[1]: m[2]
        for m in re.finditer(r'<ext uri="([^"]*)"[^>]*>(.*?)</ext>', xml[xml.rindex("<extLst>") :])
    }


def _rule_ids(xml: str) -> list[str]:
    return re.findall(r'<x14:cfRule\b[^>]*\bid="([^"]*)"', xml)


def test_the_fixture_holds_two_excel_rules() -> None:
    assert len(_rule_ids(_excel_extensions()[ext.CONDITIONAL_FORMATS])) == 2


def test_without_rules_gives_what_it_gave_before() -> None:
    ids = _rule_ids(_excel_extensions()[ext.CONDITIONAL_FORMATS])
    unknown = "{NOT-A-RULE}"
    for dead in [set(), {ids[0]}, {ids[1]}, set(ids), {unknown}, {ids[0], unknown}]:
        old, new = _excel_extensions(), _excel_extensions()
        _previous_without_rules(old, dead)
        ext.without_rules(new, dead)
        assert new == old, dead


def test_without_rules_drops_a_self_closing_rule_and_the_group_it_leaves_empty() -> None:
    xml = (
        f'<x14:conditionalFormattings><x14:conditionalFormatting xmlns:xm="{XM}">'
        '<x14:cfRule type="dataBar" id="{A}"/><xm:sqref>A1</xm:sqref></x14:conditionalFormatting>'
        '<x14:conditionalFormatting><x14:cfRule id="{B}"></x14:cfRule><xm:sqref>B1'
        "</xm:sqref></x14:conditionalFormatting></x14:conditionalFormattings>"
    )
    found = {ext.CONDITIONAL_FORMATS: xml}
    ext.without_rules(found, {"{A}"})
    assert _rule_ids(found[ext.CONDITIONAL_FORMATS]) == ["{B}"]
    assert found[ext.CONDITIONAL_FORMATS].count("xmlns:xm") == 0
    ext.without_rules(found, {"{B}"})
    assert found == {}


def _guid(number: int) -> str:
    return f"{{{number:08X}-0000-4000-8000-000000000000}}"


def _timed(work: Callable[[], object]) -> float:
    start = time.perf_counter()
    work()
    return time.perf_counter() - start


def test_many_dead_rules_are_removed_in_one_pass() -> None:
    count = 20_000
    groups = "".join(
        f'<x14:conditionalFormatting><x14:cfRule type="dataBar" id="{{{i}}}"/>'
        f"<xm:sqref>A{i}</xm:sqref></x14:conditionalFormatting>"
        for i in range(count)
    )
    found = {ext.CONDITIONAL_FORMATS: groups + groups.replace("{", "{k")}
    kept = {f"{{k{i}}}" for i in range(0, count, 2)}
    dead = {f"{{{i}}}" for i in range(count)} | {f"{{k{i}}}" for i in range(count)} - kept
    assert _timed(lambda: ext.without_rules(found, dead)) < LIMIT
    assert len(_rule_ids(found[ext.CONDITIONAL_FORMATS])) == len(kept)


class _Plain:
    edit: LineEdit | None = None

    def ranges(self, text: str, *, grow: bool = True) -> str | None:
        return text

    def formula(self, text: str) -> str:
        return text


@pytest.mark.parametrize(
    ("uri", "element"),
    [
        (ext.CONDITIONAL_FORMATS, "<x14:conditionalFormatting/>"),
        (ext.CONDITIONAL_FORMATS, '<x14:cfRule id="{A}"/>'),
        (ext.DATA_VALIDATIONS, "<x14:dataValidation/>"),
        (ext.SPARKLINES, "<x14:sparklineGroup/>"),
    ],
)
def test_self_closing_elements_do_not_make_a_scan_per_element(uri: str, element: str) -> None:
    xml = element * 30_000

    def run() -> None:
        ext.rewrite_extensions({uri: xml}, _Plain())
        ext.clip_rules({uri: xml}, uri, parse_range("A1:B2"))
        ext.forget_sheets({uri: xml}, {"Gone"})
        ext.without_rules({uri: xml}, {"{A}"})

    assert _timed(run) < LIMIT


def test_a_group_of_sparklines_without_a_data_wrapper_is_scanned_once() -> None:
    xml = (
        "<x14:sparklineGroup><x14:sparkline><xm:f>S!A1</xm:f></x14:sparkline></x14:sparklineGroup>"
    )
    found = {ext.SPARKLINES: xml * 30_000}
    assert _timed(lambda: ext.forget_sheets(found, {"Gone"})) < LIMIT


def test_many_marked_cells_are_patched_in_one_pass() -> None:
    count = 60_000
    data = "".join(f'<row r="{i}"><c r="A{i}" vm="1"><v>1</v></c></row>' for i in range(count))
    edits = {
        f"A{i}": (lambda attributes, body: (attributes + ' cm="1"', body)) for i in range(count)
    }
    result = []
    assert _timed(lambda: result.append(patch_cells(data.encode(), edits))) < LIMIT
    assert result[0].count(b' cm="1"') == count


def test_many_orphaned_slicer_caches_are_removed_in_one_pass() -> None:
    count = 20_000
    links = [
        Link(
            f"rId{i}",
            "cache",
            Part(
                f"c{i}.xml",
                consistency.SLICER_CACHE,
                f'<slicerCacheDefinition name="C{i}"/>'.encode(),
            ),
        )
        for i in range(count)
    ]
    entries = "".join(f'<x14:slicerCache r:id="rId{i}"/>' for i in range(count))
    kept: list[str] = []
    dropped = {f"C{i}" for i in range(count - 1)}
    assert (
        _timed(lambda: kept.extend(consistency.without_caches(links, [entries], dropped)[1]))
        < LIMIT
    )
    assert kept == [f'<x14:slicerCache r:id="rId{count - 1}"/>']


def _workbook_with(
    files: Path, name: str, rewrite: Callable[[str, bytes], bytes], extra: dict[str, bytes]
) -> Path:
    buffer = io.BytesIO()
    book = Workbook()
    assert book.active is not None
    book.active["A1"] = 1
    book.save(buffer)
    path = files / name
    with (
        zipfile.ZipFile(buffer) as source,
        zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as target,
    ):
        for item in source.infolist():
            target.writestr(item.filename, rewrite(item.filename, source.read(item.filename)))
        for part, data in extra.items():
            target.writestr(part, data)
    return path


@pytest.mark.filterwarnings("ignore::UserWarning")  # openpyxl says it drops the rules
async def test_a_workbook_with_many_rules_that_openpyxl_drops_is_saved_quickly(
    files: Path, call: ToolCall
) -> None:
    count = 4_000
    classic = "".join(
        f'<conditionalFormatting sqref="A{i}"><cfRule type="dataBar" priority="abc"><dataBar>'
        '<cfvo type="min"/><cfvo type="max"/><color rgb="FF00AA00"/></dataBar>'
        f'<extLst><ext uri="{RULE_URI}" xmlns:x14="{X14}"><x14:id>{_guid(i)}'
        "</x14:id></ext></extLst></cfRule></conditionalFormatting>"
        for i in range(1, count + 1)
    )
    modern = "".join(
        f'<x14:conditionalFormatting xmlns:xm="{XM}"><x14:cfRule type="dataBar" '
        f'id="{{{i:08X}-0000-4000-8000-000000000000}}"><x14:dataBar minLength="0" maxLength="100">'
        f'<x14:cfvo type="min"/><x14:cfvo type="max"/></x14:dataBar></x14:cfRule>'
        f"<xm:sqref>A{i}</xm:sqref></x14:conditionalFormatting>"
        for i in range(1, count + 1)
    )
    extension = (
        f'<extLst><ext uri="{FORMATS_URI}" xmlns:x14="{X14}"><x14:conditionalFormattings>'
        f"{modern}</x14:conditionalFormattings></ext></extLst>"
    )

    def rewrite(name: str, data: bytes) -> bytes:
        if name != "xl/worksheets/sheet1.xml":
            return data
        xml = data.decode().replace("</sheetData>", "</sheetData>" + classic)
        return xml.replace("</worksheet>", extension + "</worksheet>").encode()

    path = _workbook_with(files, "crafted.xlsx", rewrite, {})
    start = time.perf_counter()
    await call("write_range", path=str(path), sheet="Sheet", at="B2", rows=[[1]])
    assert time.perf_counter() - start < LIMIT * 2
    sheet = text(read_parts(path), "xl/worksheets/sheet1.xml")
    assert "x14:cfRule" not in sheet
    assert "<conditionalFormatting" not in sheet


async def test_a_workbook_with_many_relationships_on_a_sheet_is_saved_quickly(
    files: Path, call: ToolCall
) -> None:
    count = 8_000
    kind = "http://example.com/custom"
    rels = "".join(
        f'<Relationship Id="rIdx{i}" Type="{kind}" Target="http://example.com/{i}"'
        ' TargetMode="External"/>'
        for i in range(count)
    )
    relationships = (
        '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/'
        f'2006/relationships">{rels}</Relationships>'
    )
    path = _workbook_with(
        files,
        "rels.xlsx",
        lambda _, data: data,
        {"xl/worksheets/_rels/sheet1.xml.rels": relationships.encode()},
    )
    start = time.perf_counter()
    await call("write_range", path=str(path), sheet="Sheet", at="B2", rows=[[1]])
    assert time.perf_counter() - start < LIMIT
    saved = text(read_parts(path), "xl/worksheets/_rels/sheet1.xml.rels")
    assert saved.count(f'Type="{kind}"') == count
