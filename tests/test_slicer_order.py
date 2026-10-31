"""Excel opens a workbook with table and PivotTable slicers only when its extension lists
(``<extLst>``) are in the order Excel writes them."""

import re
from itertools import permutations
from pathlib import Path
from typing import Any

import pytest

from excel_mcp.package.extensions import in_excel_order
from tests.package_support import sheet_part, text
from tests.slicer_support import TABLE, ToolCall, add, build_book, parts_of

pytestmark = pytest.mark.anyio

PIVOT_CACHES = "{BBE1A952-AA13-448e-AADC-164F8A28A991}"
WORKBOOK_PR = "{79F54976-1DA5-4618-B147-4CDE4B953A38}"
TIMELINE_CACHES = "{D0CA8CA8-9F24-4464-BF8E-62219DCF47F9}"
TABLE_CACHES = "{46BE6895-7355-4a93-B00E-2C351335B9C9}"
PIVOT_SLICERS = "{A8765BA9-456A-4dab-B4F3-ACF838C121DE}"
TABLE_SLICERS = "{3A4CF648-6AED-40f4-86FF-DC5316D8AED3}"
TIMELINES = "{7E03D99C-DC04-49d9-9315-930204A7B6E9}"

MAKERS: dict[str, dict[str, Any]] = {
    "table": {"sheet": "Data", "target": TABLE, "field": "Region", "at": "G2"},
    "pivot": {"field": "Product"},
    "timeline": {"field": "Date", "at": "E12", "timeline": {"level": "months"}},
}


@pytest.fixture
async def book(call: ToolCall, files: Path) -> str:
    return await build_book(call, files)


def uris(parts: dict[str, bytes], part: str) -> list[str]:
    xml = text(parts, part)
    return re.findall(r'<ext uri="(\{[^}]+\})"', xml[xml.rindex("<extLst>") :])


@pytest.mark.parametrize("kinds", [p for n in (2, 3) for p in permutations(MAKERS, n)])
async def test_extensions_are_in_the_order_of_excel(
    call: ToolCall, book: str, kinds: tuple[str, ...]
) -> None:
    for kind in kinds:
        await add(call, book, **MAKERS[kind])
    parts = parts_of(book)
    expected = [
        u
        for u in (PIVOT_CACHES, WORKBOOK_PR, TIMELINE_CACHES, TABLE_CACHES)
        if u in text(parts, "xl/workbook.xml")
    ]
    assert uris(parts, "xl/workbook.xml") == expected
    sheet = uris(parts, sheet_part(parts, "Pivot"))
    assert sheet == [u for u in (PIVOT_SLICERS, TIMELINES) if u in sheet]


async def test_a_table_slicer_on_the_sheet_of_a_pivot_slicer_comes_after_it(
    call: ToolCall, book: str
) -> None:
    await add(call, book, sheet="Pivot", target=TABLE, field="Region", at="G2")
    await add(call, book, field="Product")
    await add(call, book, field="Date", at="E12", timeline={"level": "months"})
    parts = parts_of(book)
    assert uris(parts, sheet_part(parts, "Pivot")) == [PIVOT_SLICERS, TABLE_SLICERS, TIMELINES]


async def test_the_order_survives_copying_a_sheet_and_deleting_a_slicer(
    call: ToolCall, book: str
) -> None:
    await add(call, book, sheet="Data", target=TABLE, field="Region", at="G2")
    await add(call, book, field="Product")
    await add(call, book, field="Date", at="E12", timeline={"level": "months"})
    await call("copy_sheet", path=book, sheet="Pivot", new_name="Copy")
    await call("delete_slicer", path=book, sheet="Pivot", name="Product")
    parts = parts_of(book)
    assert uris(parts, "xl/workbook.xml") == [
        PIVOT_CACHES,
        WORKBOOK_PR,
        TIMELINE_CACHES,
        TABLE_CACHES,
    ]


def test_only_the_entries_that_excel_orders_move() -> None:
    entries = {TABLE_CACHES: "a", "{OTHER}": "b", PIVOT_CACHES: "c", "{LAST}": "d"}
    assert list(in_excel_order(entries)) == [PIVOT_CACHES, "{OTHER}", TABLE_CACHES, "{LAST}"]
    assert in_excel_order(entries)[PIVOT_CACHES] == "c"
