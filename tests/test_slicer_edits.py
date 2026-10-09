"""Slicers and timelines when the workbook around them changes: deleting, copying, editing."""

import re
import zipfile
from pathlib import Path
from typing import Any

import pytest

from tests.package_support import read_parts, text
from tests.slicer_support import (
    PIVOT,
    TABLE,
    CallError,
    ToolCall,
    add,
    build_book,
    named,
    parts_of,
    pivot_cells,
)

pytestmark = pytest.mark.anyio


@pytest.fixture
async def book(call: ToolCall, files: Path) -> str:
    return await build_book(call, files)


def cache_names(path: str) -> list[str]:
    parts = parts_of(path)
    found = []
    for xml in named(parts, "xl/slicerCaches/") + named(parts, "xl/timelineCaches/"):
        found.append(re.search(r'name="([^"]+)"', xml.split("?>")[1])[1])  # type: ignore[index]
    return sorted(found)


def grand_total(path: str, sheet: str = "Pivot") -> Any:
    rows = pivot_cells(path, sheet)
    return next(total for label, total in rows if label == "Grand Total")


async def slicers_of(call: ToolCall, path: str, sheet: str) -> list[dict[str, Any]]:
    return (await call("describe_sheet", path=path, sheet=sheet)).get("slicers", [])


async def test_deleting_a_slicer_removes_it_and_its_cache(call: ToolCall, book: str) -> None:
    await add(call, book, field="Product")
    await call("delete_slicer", path=book, sheet="Pivot", name="Product")
    parts = parts_of(book)
    assert not [n for n in parts if "slicer" in n.lower()]
    assert "Slicer_Product" not in text(parts, "xl/workbook.xml")
    assert "slicerCache" not in text(parts, "xl/workbook.xml")
    assert await slicers_of(call, book, "Pivot") == []


async def test_deleting_keeps_the_filter_of_a_row_field_and_of_a_table(
    call: ToolCall, book: str
) -> None:
    await add(call, book, field="Product", selected_items=["Apples"])
    await add(call, book, sheet="Data", target=TABLE, field="Region", selected_items=["North"])
    await call("delete_slicer", path=book, sheet="Pivot", name="Product")
    await call("delete_slicer", path=book, sheet="Data", name="Region")
    assert pivot_cells(book)[:3] == [
        ["Product", "Sum of Amount"],
        ["Apples", 100],
        ["Grand Total", 100],
    ]
    parts = parts_of(book)
    assert 'h="1"' in text(parts, "xl/pivotTables/pivotTable1.xml")
    assert "<filter val" in text(parts, "xl/tables/table1.xml")


async def test_deleting_clears_what_only_the_slicer_kept(call: ToolCall, book: str) -> None:
    await add(call, book, field="Region", selected_items=["North"])
    await add(
        call, book, field="Date", at="H3", timeline={"start": "2025-01-01", "end": "2025-02-28"}
    )
    assert grand_total(book) == 10
    await call("delete_slicer", path=book, sheet="Pivot", name="Date")
    assert grand_total(book) == 40
    await call("delete_slicer", path=book, sheet="Pivot", name="Region")
    assert grand_total(book) == 360
    assert 'h="1"' not in text(parts_of(book), "xl/pivotTables/pivotTable1.xml")
    assert "dateBetween" not in text(parts_of(book), "xl/pivotTables/pivotTable1.xml")


async def test_deleting_an_unknown_slicer_lists_the_known_ones(
    call: ToolCall, call_error: CallError, book: str
) -> None:
    await add(call, book, field="Product")
    error = await call_error("delete_slicer", path=book, sheet="Pivot", name="Nope")
    assert "has no slicer or timeline 'Nope'. Available: 'Product'" in error


async def test_a_pivot_table_with_slicers_cannot_be_deleted(
    call: ToolCall, call_error: CallError, book: str
) -> None:
    await add(call, book, field="Product")
    error = await call_error("delete_pivot_table", path=book, sheet="Pivot", name="PivotSales")
    assert "Slicer_Product" in error
    await call("delete_slicer", path=book, sheet="Pivot", name="Product")
    await call("delete_pivot_table", path=book, sheet="Pivot", name="PivotSales")


async def test_copying_a_sheet_copies_its_slicers_with_their_own_caches(
    call: ToolCall, book: str
) -> None:
    await add(call, book, field="Product", selected_items=["Apples", "Pears"])
    await add(call, book, field="Date", at="H3", timeline={"level": "years"})
    await add(call, book, sheet="Data", target=TABLE, field="Region", selected_items=["East"])
    await call("copy_sheet", path=book, sheet="Pivot", new_name="Pivot 2")
    await call("copy_sheet", path=book, sheet="Data", new_name="Data 2")
    copied = await slicers_of(call, book, "Pivot 2")
    assert [(s["name"], s["kind"], s["target"]) for s in copied] == [
        ("Product 1", "pivot", "Pivot 2!PivotSales"),
        ("Date 1", "timeline", "Pivot 2!PivotSales"),
    ]
    assert copied[0]["selected_items"] == ["Apples", "Pears"]
    assert pivot_cells(book, "Pivot 2")[:4] == pivot_cells(book)[:4]
    table = (await slicers_of(call, book, "Data 2"))[0]
    assert (table["name"], table["target"], table["selected_items"]) == (
        "Region 1",
        "Data 2!Sales2",
        ["East"],
    )
    assert cache_names(book) == [
        "NativeTimeline_Date",
        "NativeTimeline_Date1",
        "Slicer_Product",
        "Slicer_Product1",
        "Slicer_Region",
        "Slicer_Region1",
    ]
    workbook = text(parts_of(book), "xl/workbook.xml")
    assert '<definedName name="Slicer_Product1">#N/A</definedName>' in workbook


async def test_a_copied_slicer_filters_only_the_copy(call: ToolCall, book: str) -> None:
    await add(call, book, field="Region", selected_items=["North"])
    await call("copy_sheet", path=book, sheet="Pivot", new_name="Pivot 2")
    await call("delete_slicer", path=book, sheet="Pivot 2", name="Region 1")
    assert grand_total(book) == 40
    assert grand_total(book, "Pivot 2") == 360
    assert cache_names(book) == ["Slicer_Region"]


async def test_copying_the_sheet_of_the_pivot_tables_joins_the_slicers_elsewhere(
    call: ToolCall, book: str
) -> None:
    await call("create_sheet", path=book, new_name="Dashboard")
    await add(call, book, sheet="Dashboard", field="Region", selected_items=["East"])
    await call("copy_sheet", path=book, sheet="Pivot", new_name="Pivot 2")
    info = await slicers_of(call, book, "Dashboard")
    assert info[0]["target"] == "Pivot!PivotSales, Pivot 2!PivotSales"
    assert cache_names(book) == ["Slicer_Region"]
    assert grand_total(book, "Pivot 2") == 120


async def test_copying_the_sheet_of_the_slicers_shares_the_cache(call: ToolCall, book: str) -> None:
    await call("create_sheet", path=book, new_name="Dashboard")
    await add(call, book, sheet="Dashboard", field="Region", selected_items=["East"])
    await call("copy_sheet", path=book, sheet="Dashboard", new_name="Dashboard 2")
    copied = await slicers_of(call, book, "Dashboard 2")
    assert [(s["name"], s["target"]) for s in copied] == [("Region 1", "Pivot!PivotSales")]
    assert cache_names(book) == ["Slicer_Region"]


async def test_one_slicer_filters_the_pivot_tables_that_share_a_cache(
    call: ToolCall, book: str
) -> None:
    await call("copy_sheet", path=book, sheet="Pivot", new_name="Pivot 2")
    await add(
        call,
        book,
        field="Region",
        selected_items=["East", "North"],
        connect=[{"sheet": "Pivot 2", "name": "PivotSales"}],
    )
    expected = [["Product", "Sum of Amount"], ["Apples", 80], ["Kiwis", 50], ["Pears", 30]]
    assert pivot_cells(book)[:4] == expected
    assert pivot_cells(book, "Pivot 2")[:4] == expected
    cache = named(parts_of(book), "xl/slicerCaches/")[0]
    assert cache.count("<pivotTable ") == 2
    await call("delete_slicer", path=book, sheet="Pivot", name="Region")
    assert grand_total(book) == 360 and pivot_cells(book, "Pivot 2")[4][1] == 360


async def test_connecting_needs_a_shared_cache(
    call: ToolCall, call_error: CallError, book: str
) -> None:
    await call(
        "create_pivot_table",
        path=book,
        source="Data!A1:D9",
        row_fields=["Region"],
        value_fields=[{"field": "Amount"}],
        sheet="Pivot",
        at="H3",
        name="Other",
    )
    error = await call_error(
        "add_slicer",
        path=book,
        sheet="Pivot",
        target=PIVOT,
        field="Region",
        at="E3",
        connect=[{"sheet": "Pivot", "name": "Other"}],
    )
    assert "shares 'PivotSales''s data cache" in error


async def test_a_slicer_moves_with_the_cells_it_sits_on(call: ToolCall, book: str) -> None:
    await add(call, book, field="Product", at="E5")
    await call("insert_rows_or_columns", path=book, sheet="Pivot", axis="rows", start=1, count=2)
    await call("insert_rows_or_columns", path=book, sheet="Pivot", axis="columns", start=1, count=1)
    assert (await slicers_of(call, book, "Pivot"))[0]["range"].startswith("F7:")


async def test_deleting_a_sliced_column_removes_the_table_slicer(call: ToolCall, book: str) -> None:
    await add(call, book, sheet="Data", target=TABLE, field="Region", selected_items=["East"])
    await add(call, book, sheet="Data", target=TABLE, field="Product", at="G20")
    await call("delete_rows_or_columns", path=book, sheet="Data", axis="columns", start=1, count=1)
    info = await slicers_of(call, book, "Data")
    assert [s["name"] for s in info] == ["Product"]
    assert cache_names(book) == ["Slicer_Product"]
    parts = parts_of(book)
    assert '<x15:tableSlicerCache tableId="1" column="2"' in named(parts, "xl/slicerCaches/")[0]


async def test_deleting_the_rows_of_a_table_removes_its_slicers(call: ToolCall, book: str) -> None:
    await call("create_sheet", path=book, new_name="Lists")
    await call("write_range", path=book, sheet="Lists", at="A1", rows=[["Tag"], ["a"], ["b"]])
    await call("create_table", path=book, sheet="Lists", range="A1:A3", name="Tags")
    await add(
        call, book, sheet="Lists", target={"sheet": "Lists", "name": "Tags"}, field="Tag", at="D1"
    )
    await call("delete_rows_or_columns", path=book, sheet="Lists", axis="rows", start=1, count=3)
    assert cache_names(book) == []
    assert not [n for n in parts_of(book) if "slicer" in n.lower()]


async def test_deleting_a_sheet_removes_the_slicers_on_it(call: ToolCall, book: str) -> None:
    await call("create_sheet", path=book, new_name="Dashboard")
    await add(call, book, sheet="Dashboard", field="Product")
    await add(call, book, sheet="Data", target=TABLE, field="Region")
    await call("delete_slicer", path=book, sheet="Dashboard", name="Product")
    await call("delete_sheet", path=book, sheet="Dashboard")
    assert cache_names(book) == ["Slicer_Region"]


async def test_a_table_with_slicers_on_other_sheets_cannot_lose_its_sheet(
    call: ToolCall, call_error: CallError, book: str
) -> None:
    await call("create_sheet", path=book, new_name="Dashboard")
    await add(call, book, sheet="Dashboard", target=TABLE, field="Region")
    error = await call_error("delete_sheet", path=book, sheet="Data")
    assert "Slicer_Region" in error


def _without_refresh_mark(path: str) -> None:
    """Make the PivotTable look refreshed by Excel, as if the file had been through it."""
    parts = read_parts(Path(path))
    name = "xl/pivotCache/pivotCacheDefinition1.xml"
    parts[name] = parts[name].replace(b'refreshedBy="excel-mcp-server"', b'refreshedBy="Someone"')
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        for part, data in parts.items():
            archive.writestr(part, data)


async def test_a_pivot_table_excel_made_is_filtered_for_excel_to_recalculate(
    call: ToolCall, book: str
) -> None:
    _without_refresh_mark(book)
    message = await add(call, book, field="Region", selected_items=["East"])
    assert "Excel recalculates them when the file is opened" in message["note"]
    assert grand_total(book) == 360  # the cells keep their figures
    parts = parts_of(book)
    assert 'h="1"' in text(parts, "xl/pivotTables/pivotTable1.xml")
    assert 'refreshOnLoad="1"' in text(parts, "xl/pivotCache/pivotCacheDefinition1.xml")
    info = await slicers_of(call, book, "Pivot")
    assert info[0]["selected_items"] == ["East"]
    await add(
        call, book, field="Date", at="H3", timeline={"start": "2025-01-01", "end": "2025-02-28"}
    )
    assert 'type="dateBetween"' in text(parts_of(book), "xl/pivotTables/pivotTable1.xml")
    await call("delete_slicer", path=book, sheet="Pivot", name="Date")
    assert "dateBetween" not in text(parts_of(book), "xl/pivotTables/pivotTable1.xml")
