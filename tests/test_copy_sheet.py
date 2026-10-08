from pathlib import Path

import pytest
from openpyxl import Workbook, load_workbook
from openpyxl.worksheet.worksheet import Worksheet
from PIL import Image

from excel_mcp.operations.images import list_images
from excel_mcp.operations.pivot_index import sheet_pivots
from tests.conftest import ToolCall

pytestmark = pytest.mark.anyio

BOOK = {"path": "sales.xlsx"}


@pytest.fixture
def picture(files: Path) -> Path:
    path = files / "logo.png"
    Image.new("RGB", (200, 100), "red").save(path)
    return path


def _sheet(path: Path, name: str) -> Worksheet:
    sheet = load_workbook(path)[name]
    assert isinstance(sheet, Worksheet)
    return sheet


async def _decorate_data_sheet(call: ToolCall) -> None:
    await call("write_range", **BOOK, sheet="Data", start_cell="C6", rows=[["=SUM(Sales[Units])"]])
    await call("write_range", **BOOK, sheet="Data", start_cell="D6", rows=[["=SUM(Data!D2:D5)"]])
    await call("create_table", **BOOK, sheet="Data", range="A1:D5", name="Sales")
    await call(
        "add_data_validation",
        **BOOK,
        sheet="Data",
        range="B2:B5",
        rule={"type": "list", "options": ["Apples", "Pears"]},
    )
    await call(
        "add_conditional_format",
        **BOOK,
        sheet="Data",
        range="C2:C5",
        rule={"type": "cell_value", "operator": "greaterThan", "values": ["6"]},
    )
    await call("set_note", **BOOK, sheet="Data", cell="A2", text="check")
    await call("insert_image", **BOOK, sheet="Data", image_path="logo.png", cell="G2")
    await call(
        "create_chart",
        **BOOK,
        sheet="Data",
        data_range="A1:C5",
        chart_type="column",
        anchor_cell="G10",
    )
    await call(
        "create_chart",
        **BOOK,
        sheet="Data",
        data_sheet="Report",
        data_range="A1:B3",
        chart_type="line",
        anchor_cell="G30",
    )
    await call("set_defined_name", **BOOK, name="Local", refers_to="Data!$A$1:$A$3", sheet="Data")
    await call("merge_cells", **BOOK, sheet="Data", range="A8:C8")
    await call(
        "set_sheet_layout",
        **BOOK,
        sheet="Data",
        layout={
            "column_widths": {"A": 22},
            "freeze_panes": "A2",
            "rows": [{"span": "3", "action": "hide"}],
            "print_setup": {"orientation": "landscape", "print_area": "A1:D6", "title_rows": "1:1"},
            "protection": {"enabled": True},
        },
    )
    await call(
        "create_pivot_table",
        **BOOK,
        source_sheet="Data",
        source_range="A1:D5",
        rows=["Region"],
        values=[{"field": "Units"}],
        target_sheet="Data",
        target_cell="N2",
    )


async def test_copy_includes_everything_the_server_supports(
    call: ToolCall, sample: Path, picture: Path
) -> None:
    await call("write_range", **BOOK, sheet="Report", start_cell="A1", rows=[["k", "v"], ["a", 1]])
    await _decorate_data_sheet(call)
    await call("copy_sheet", **BOOK, sheet="Data", new_name="Copy")

    assert load_workbook(sample).sheetnames == ["Data", "Report", "Copy"]
    original, copied = _sheet(sample, "Data"), _sheet(sample, "Copy")
    assert copied["A2"].value == "North"
    assert copied["A2"].comment is not None
    assert copied["C6"].value == "=SUM(Sales2[Units])"
    assert copied["D6"].value == "=SUM('Copy'!D2:D5)"
    assert original["C6"].value == "=SUM(Sales[Units])"
    assert list(copied.tables) == ["Sales2"]
    assert list(original.tables) == ["Sales"]
    assert [str(v.sqref) for v in copied.data_validations.dataValidation] == ["B2:B5"]
    assert [str(entry.sqref) for entry in copied.conditional_formatting] == ["C2:C5"]
    assert len(list_images(copied)) == 1
    assert [str(r) for r in copied.merged_cells.ranges] == ["A8:C8"]
    assert copied.column_dimensions["A"].width == 22
    assert copied.row_dimensions[3].hidden
    assert copied.freeze_panes == "A2"
    assert copied.protection.sheet
    assert copied.page_setup.orientation == "landscape"
    assert copied.print_area == "'Copy'!$A$1:$D$6"
    assert copied.print_title_rows == "$1:$1"
    assert copied.defined_names["Local"].attr_text == "'Copy'!$A$1:$A$3"
    assert not any(view.tabSelected for view in copied.views.sheetView)
    assert [pivot.name for pivot in sheet_pivots(copied)] == [
        pivot.name for pivot in sheet_pivots(original)
    ]


async def test_copied_charts_point_at_the_copy_only_when_their_data_is_on_the_source(
    call: ToolCall, sample: Path, picture: Path
) -> None:
    await call("write_range", **BOOK, sheet="Report", start_cell="A1", rows=[["k", "v"], ["a", 1]])
    await _decorate_data_sheet(call)
    await call("copy_sheet", **BOOK, sheet="Data", new_name="Copy")

    own, other = _sheet(sample, "Copy")._charts  # pyright: ignore[reportAttributeAccessIssue]
    own_refs = {s.val.numRef.f for s in own.series} | {s.cat.numRef.f for s in own.series}
    assert own_refs and all(ref.startswith("'Copy'!") for ref in own_refs)
    assert {s.val.numRef.f for s in other.series} == {"'Report'!$B$2:$B$3"}
    original = _sheet(sample, "Data")._charts[0]  # pyright: ignore[reportAttributeAccessIssue]
    assert all(s.val.numRef.f.startswith("'Data'!") for s in original.series)


async def test_copy_keeps_the_source_untouched_and_names_unique(
    call: ToolCall, sample: Path, picture: Path
) -> None:
    await _decorate_data_sheet(call)
    await call("copy_sheet", **BOOK, sheet="Data", new_name="Copy")
    await call("copy_sheet", **BOOK, sheet="Data", new_name="Copy 2")
    assert list(_sheet(sample, "Copy").tables) == ["Sales2"]
    assert list(_sheet(sample, "Copy 2").tables) == ["Sales3"]
    assert _sheet(sample, "Copy 2")["C6"].value == "=SUM(Sales3[Units])"


async def test_copy_refuses_unsafe_formulas(call_error: ToolCall, files: Path) -> None:
    workbook = Workbook()
    workbook.worksheets[0]["A1"] = '=INDIRECT("B1")'
    workbook.save(files / "unsafe.xlsx")
    message = await call_error("copy_sheet", path="unsafe.xlsx", sheet="Sheet", new_name="Copy")
    assert "INDIRECT" in message
