"""Edits that change what preserved content refers to: Excel's rules, or a refusal."""

import re
import zipfile
from pathlib import Path

import pytest
from mcp import Client
from openpyxl import Workbook
from openpyxl.formatting.formatting import ConditionalFormattingList

from excel_mcp.config import Limits, Settings
from excel_mcp.paths import PathPolicy
from excel_mcp.server import create_server
from excel_mcp.workspace import Workspace
from tests.conftest import ToolCall
from tests.package_support import (
    assert_package_is_consistent,
    copy_fixture,
    read_parts,
    sheet_part,
    text,
)

pytestmark = pytest.mark.anyio

BOOK = {"path": "book.xlsx"}


async def test_preserved_parts_are_copied_byte_for_byte(call: ToolCall, files: Path) -> None:
    for fixture, prefixes in (
        ("excel_chartex.xlsx", ("xl/charts/chartEx", "xl/charts/style2", "xl/charts/colors3")),
        ("excel_slicers.xlsx", ("xl/slicers/", "xl/timelines/", "xl/slicerCaches/")),
    ):
        source = copy_fixture(files, fixture)
        before = {n: d for n, d in read_parts(source).items() if n.startswith(prefixes)}
        await call("write_range", **BOOK, sheet="Data", at="Z1", rows=[[1]])
        after = read_parts(files / "book.xlsx")
        assert before
        for name, data in before.items():
            if "slicerCaches/" in name:
                continue  # sheet numbers inside are renumbered, see the slicer tests
            assert after[name] == data, name


async def test_deleting_a_sheet_removes_slicers_nothing_uses(call: ToolCall, files: Path) -> None:
    copy_fixture(files, "excel_slicers.xlsx")
    await call("delete_sheet", **BOOK, sheet="Data")
    parts = read_parts(files / "book.xlsx")
    workbook = text(parts, "xl/workbook.xml")

    assert sorted(n for n in parts if "slicerCaches/" in n or "timelineCaches/" in n) != []
    names = re.findall(r'<definedName name="([^"]*)"', workbook)
    assert sorted(names) == ["NativeTimeline_Date", "Slicer_Product"]
    assert not [n for n in parts if n.startswith("xl/tables/")]
    assert "tableSlicerCache" not in b"".join(parts.values()).decode("utf-8", "ignore")
    assert workbook.count("<x14:slicerCache ") == 1
    assert_package_is_consistent(parts)


async def test_deleting_a_sheet_keeps_content_that_remains_valid(
    call: ToolCall, files: Path
) -> None:
    copy_fixture(files, "excel_sparklines.xlsx")
    await call("delete_sheet", **BOOK, sheet="Lists")
    parts = read_parts(files / "book.xlsx")
    sheet = text(parts, sheet_part(parts, "Data"))

    assert sheet.count("<x14:sparkline>") == 18  # they read the Data sheet
    assert "<xm:f>#REF!</xm:f>" in sheet  # the validation listed cells of Lists
    assert_package_is_consistent(parts)


async def test_a_sheet_with_pivot_tables_that_slicers_elsewhere_use_is_kept(
    call: ToolCall, call_error: object, files: Path
) -> None:
    source = copy_fixture(files, "excel_slicers.xlsx")
    before = source.read_bytes()
    server = create_server(Settings(allowed_dirs=[files]))
    async with Client(server) as client:
        for tool, arguments in (
            ("delete_sheet", {"sheet": "Pivot"}),
            ("delete_pivot_table", {"sheet": "Pivot", "name": "PivotSales"}),
        ):
            result = await client.call_tool(tool, {**BOOK, **arguments})
            assert result.is_error
            message = str(result.content[0])
            assert "connected" in message and "Delete those slicers in Excel first" in message
    assert source.read_bytes() == before


async def test_a_rule_that_lost_its_main_rule_loses_its_extension(files: Path) -> None:
    copy_fixture(files, "excel_sparklines.xlsx")
    workspace = Workspace(PathPolicy([files]), Limits(), allow_macro_workbooks=False)

    with workspace.edit("book.xlsx") as workbook:
        workbook["Data"].conditional_formatting = ConditionalFormattingList()
    parts = read_parts(files / "book.xlsx")
    sheet = text(parts, sheet_part(parts, "Data"))

    assert '<x14:cfRule type="dataBar"' not in sheet
    assert '<x14:cfRule type="iconSet"' in sheet  # Excel stores this one only as an extension
    assert sheet.count("<x14:sparklineGroup ") == 3


async def test_preserved_content_beyond_the_size_limit_is_refused(files: Path) -> None:
    path = files / "book.xlsx"
    Workbook().save(path)
    with zipfile.ZipFile(path) as archive:
        parts = {name: archive.read(name) for name in archive.namelist()}
    parts["customXml/item1.xml"] = b"<a>" + b"x" * 5_000_000 + b"</a>"  # compresses to 5 KB
    parts["xl/_rels/workbook.xml.rels"] = parts["xl/_rels/workbook.xml.rels"].replace(
        b"</Relationships>",
        b'<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
        b'relationships/customXml" Target="../customXml/item1.xml"/></Relationships>',
    )
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, data in parts.items():
            archive.writestr(name, data)
    assert path.stat().st_size < 100_000
    limits = Limits(max_file_bytes=1_000_000, max_compression_ratio=10**6)
    server = create_server(Settings(allowed_dirs=[files], limits=limits))

    async with Client(server) as client:
        result = await client.call_tool(
            "write_range", {**BOOK, "sheet": "Sheet", "at": "A1", "rows": [[1]]}
        )

    assert result.is_error
    assert "bytes this server can carry along" in str(result.content[0])
