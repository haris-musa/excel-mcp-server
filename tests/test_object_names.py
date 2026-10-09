"""Charts and images are selected by name, and keep the names Excel gave them."""

import re
import zipfile
from pathlib import Path

import pytest
from PIL import Image

from tests.conftest import ToolCall

pytestmark = pytest.mark.anyio

BOOK = {"path": "sales.xlsx", "sheet": "Report"}


@pytest.fixture
def picture(files: Path) -> Path:
    path = files / "logo.png"
    Image.new("RGB", (40, 20), "red").save(path)
    return path


async def chart(call: ToolCall, at: str, **extra: object) -> dict[str, str]:
    return await call(
        "create_chart", **BOOK, chart_type="column", source="Data!B1:C5", at=at, **extra
    )


def drawing(path: Path) -> str:
    with zipfile.ZipFile(path) as archive:
        return archive.read("xl/drawings/drawing1.xml").decode("utf-8")


def rename_in_drawing(path: Path, old: str, new: str) -> None:
    """What Excel does when a chart or picture is renamed in its Name Box."""
    with zipfile.ZipFile(path) as archive:
        parts = {name: archive.read(name) for name in archive.namelist()}
    parts["xl/drawings/drawing1.xml"] = (
        parts["xl/drawings/drawing1.xml"].decode("utf-8").replace(f'name="{old}"', f'name="{new}"')
    ).encode("utf-8")
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, data in parts.items():
            archive.writestr(name, data)


async def test_new_charts_are_named_like_excel_names_them(call: ToolCall, sample: Path) -> None:
    first = await chart(call, "B2")
    second = await chart(call, "B20", name="Totals")
    third = await chart(call, "B40")
    assert [item["name"] for item in (first, second, third)] == ["Chart 1", "Totals", "Chart 2"]
    names = re.findall(r'<(?:\w+:)?cNvPr id="\d+" name="([^"]*)"', drawing(sample))
    assert names == ["Chart 1", "Totals", "Chart 2"]


async def test_names_do_not_move_when_other_charts_are_deleted(
    call: ToolCall, sample: Path
) -> None:
    for at in ("B2", "B20", "B40"):
        await chart(call, at)
    await call("delete_chart", **BOOK, name="Chart 2")
    details = await call("describe_sheet", **BOOK)
    assert [(item["name"], item["range"].split(":")[0]) for item in details["charts"]] == [
        ("Chart 1", "B2"),
        ("Chart 3", "B40"),
    ]
    await call("delete_chart", **BOOK, name="chart 3")
    details = await call("describe_sheet", **BOOK)
    assert [item["name"] for item in details["charts"]] == ["Chart 1"]


async def test_chart_names_must_be_unique_and_not_blank(
    call: ToolCall, call_error: ToolCall, sample: Path
) -> None:
    await chart(call, "B2", name="Revenue")
    assert "already exists" in await call_error(
        "create_chart", **BOOK, chart_type="line", source="Data!B1:C5", at="B20", name="revenue"
    )
    assert "1 to 255" in await call_error(
        "create_chart", **BOOK, chart_type="line", source="Data!B1:C5", at="B20", name=" "
    )
    details = await call("describe_sheet", **BOOK)
    assert len(details["charts"]) == 1


async def test_charts_and_images_share_one_name_space(
    call: ToolCall, call_error: ToolCall, sample: Path, picture: Path
) -> None:
    await chart(call, "B2", name="Logo")
    arguments = {**BOOK, "image_path": "logo.png", "at": "H2"}
    assert "already exists" in await call_error("insert_image", **arguments, name="Logo")
    placed = await call("insert_image", **arguments)
    assert placed["name"] == "Picture 1"


async def test_names_from_excel_survive_edits(call: ToolCall, sample: Path, picture: Path) -> None:
    await chart(call, "B2")
    await call("insert_image", **BOOK, image_path="logo.png", at="H20")
    rename_in_drawing(sample, "Chart 1", "Sales chart")
    rename_in_drawing(sample, "Picture 1", "Company logo")

    details = await call("describe_sheet", **BOOK)
    assert [item["name"] for item in details["charts"]] == ["Sales chart"]
    assert [item["name"] for item in details["images"]] == ["Company logo"]

    await call("write_range", **BOOK, at="A1", rows=[["edited"]])
    assert re.findall(r'<(?:\w+:)?cNvPr id="\d+" name="([^"]*)"', drawing(sample)) == [
        "Sales chart",
        "Company logo",
    ]
    await call("delete_chart", **BOOK, name="Sales chart")
    await call("delete_image", **BOOK, name="Company logo")
    assert "xl/drawings/drawing1.xml" not in zipfile.ZipFile(sample).namelist()


async def test_duplicate_names_in_a_file_are_made_unique(
    call: ToolCall, sample: Path, picture: Path
) -> None:
    await chart(call, "B2")
    await call("insert_image", **BOOK, image_path="logo.png", at="H20")
    rename_in_drawing(sample, "Picture 1", "Chart 1")

    details = await call("describe_sheet", **BOOK)
    assert [item["name"] for item in details["charts"]] == ["Chart 1"]
    assert [item["name"] for item in details["images"]] == ["Picture 1"]


async def test_the_range_of_a_chart_follows_the_column_widths(call: ToolCall, sample: Path) -> None:
    await call(
        "set_sheet_layout", **BOOK, layout={"column_widths_chars": {"B": 60, "C": 60, "D": 60}}
    )
    made = await chart(call, "B2")
    assert made["range"] == "B2:C16"
    details = await call("describe_sheet", **BOOK)
    assert details["charts"][0]["range"] == "B2:C16"


async def test_results_tell_where_the_object_went(
    call: ToolCall, sample: Path, picture: Path
) -> None:
    placed = await call(
        "insert_image", **BOOK, image_path="logo.png", at="B2", width_cm=4, name="Logo"
    )
    assert placed == {"sheet": "Report", "name": "Logo", "range": "B2:D5"}


async def test_describe_workbook_counts_the_objects_of_each_sheet(
    call: ToolCall, sample: Path, picture: Path
) -> None:
    await call("create_table", path="sales.xlsx", sheet="Data", range="A1:D5")
    await chart(call, "H2")
    await chart(call, "H20")
    await call("insert_image", **BOOK, image_path="logo.png", at="H40")
    await call(
        "create_pivot_table",
        path="sales.xlsx",
        sheet="Report",
        at="A3",
        source="Data!A1:D5",
        row_fields=["Region"],
        value_fields=[{"field": "Units"}],
    )
    await call(
        "add_slicer",
        **BOOK,
        target={"sheet": "Report", "name": "PivotTable1"},
        field="Region",
        at="A20",
    )
    await call("set_sheet_layout", path="sales.xlsx", sheet="Data", layout={"visibility": "hidden"})

    sheets = (await call("describe_workbook", path="sales.xlsx"))["sheets"]

    assert [(item["name"], item.get("visibility"), item.get("objects")) for item in sheets] == [
        ("Data", "hidden", {"tables": 1}),
        ("Report", None, {"charts": 2, "pivot_tables": 1, "slicers": 1, "images": 1}),
    ]
