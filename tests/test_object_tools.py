from pathlib import Path

import pytest
from openpyxl import load_workbook

from tests.conftest import ToolCall

pytestmark = pytest.mark.anyio


async def test_create_table(call: ToolCall, call_error: ToolCall, sample: Path) -> None:
    message = await call("create_table", path="sales.xlsx", sheet="Data", range="A1:D5")
    assert "Table1" in message
    assert load_workbook(sample)["Data"].tables["Table1"].ref == "A1:D5"
    assert "already used" in await call_error(
        "create_table", path="sales.xlsx", sheet="Report", range="A1:A2", name="table1"
    )


async def test_create_table_requires_text_headers(call_error: ToolCall, sample: Path) -> None:
    message = await call_error("create_table", path="sales.xlsx", sheet="Data", range="C2:D5")
    assert "header" in message


@pytest.mark.parametrize("chart_type", ["column", "bar", "line", "area", "pie", "scatter"])
async def test_create_chart(call: ToolCall, sample: Path, chart_type: str) -> None:
    await call(
        "create_chart",
        path="sales.xlsx",
        sheet="Report",
        data_sheet="Data",
        data_range="B1:C5",
        chart_type=chart_type,
        anchor_cell="B2",
        options={"title": "Units", "x_axis_title": "Product"},
    )
    details = await call("describe_sheet", path="sales.xlsx", sheet="Report")
    assert len(details["charts"]) == 1


async def test_chart_needs_labels_and_a_series(call_error: ToolCall, sample: Path) -> None:
    message = await call_error(
        "create_chart",
        path="sales.xlsx",
        sheet="Data",
        data_range="C1:C5",
        chart_type="line",
        anchor_cell="F2",
    )
    assert "header row" in message


async def test_create_summary_table(call: ToolCall, sample: Path) -> None:
    await call(
        "create_summary_table",
        path="sales.xlsx",
        sheet="Data",
        source_range="A1:D5",
        group_by=["region"],
        values=[{"field": "Units"}, {"field": "Price", "aggregation": "average"}],
        target_sheet="Report",
        target_cell="B3",
    )
    sheet = load_workbook(sample)["Report"]
    rows = [[cell.value for cell in row] for row in sheet["B3:D5"]]
    assert rows == [
        ["region", "Units (sum)", "Price (average)"],
        ["North", 17, 1.75],
        ["South", 8, 1.75],
    ]


async def test_summary_reports_unknown_fields(call_error: ToolCall, sample: Path) -> None:
    message = await call_error(
        "create_summary_table",
        path="sales.xlsx",
        sheet="Data",
        source_range="A1:D5",
        group_by=["Country"],
        values=[{"field": "Units"}],
        target_sheet="Report",
    )
    assert "Available fields" in message
