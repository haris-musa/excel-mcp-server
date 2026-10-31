"""Chart and layout requests are sized before anything is built."""

import time
from collections.abc import AsyncIterator
from pathlib import Path

import pytest
from mcp import Client

from excel_mcp.config import Limits, Settings
from excel_mcp.server import create_server
from tests.conftest import ToolCall, error_text

pytestmark = pytest.mark.anyio

BIG = "A1:XFD1048576"


@pytest.fixture
async def small(files: Path) -> AsyncIterator[Client]:
    settings = Settings(allowed_dirs=[files], limits=Limits(max_cells=200))
    async with Client(create_server(settings)) as client:
        yield client


async def _chart(client: Client, **arguments: object) -> tuple[bool, str, float]:
    started = time.monotonic()
    result = await client.call_tool(
        "create_chart", {"path": "sales.xlsx", "sheet": "Data", "at": "H2", **arguments}
    )
    text = error_text(result)
    return result.is_error, text, time.monotonic() - started


@pytest.mark.parametrize(
    ("chart_type", "series_in"),
    [("column", "rows"), ("column", "columns"), ("waterfall", "columns")],
)
async def test_a_chart_range_is_sized_before_series_are_built(
    small: Client, sample: Path, series_in: str, chart_type: str
) -> None:
    failed, message, seconds = await _chart(
        small, chart_type=chart_type, source=BIG, series_in=series_in
    )
    assert failed
    assert "at most 200" in message
    assert seconds < 5


async def test_explicit_series_and_categories_are_sized_too(small: Client, sample: Path) -> None:
    for arguments in (
        {"series": [{"values": "B1:B1048576"}]},
        {"series": [{"values": "B1:B5", "categories": "A1:A1048576"}]},
        {"series": [{"values": "B1:B5", "sizes": "A1:A1048576"}], "categories": "A1:A5"},
    ):
        failed, message, _ = await _chart(small, chart_type="column", **arguments)
        assert failed
        assert "at most 200" in message


async def test_the_series_count_is_checked_from_the_range(files: Path, sample: Path) -> None:
    settings = Settings(allowed_dirs=[files], limits=Limits(max_cells=100_000))
    async with Client(create_server(settings)) as client:
        failed, message, seconds = await _chart(
            client, chart_type="column", source="A1:C300", series_in="rows"
        )
    assert failed
    assert "at most 255 series, got 299" in message
    assert seconds < 5


async def test_a_chart_at_the_series_limit_is_created(files: Path, sample: Path) -> None:
    settings = Settings(allowed_dirs=[files], limits=Limits(max_cells=100_000))
    async with Client(create_server(settings)) as client:
        failed, message, _ = await _chart(
            client, chart_type="line", source="A1:C256", series_in="rows"
        )
    assert not failed, message


async def test_row_actions_share_the_row_limit(
    call: ToolCall, call_error: ToolCall, sample: Path
) -> None:
    actions = [
        {"span": f"{start}:{start + 4_999}", "action": "hide"} for start in (1, 5_001, 10_001)
    ]
    message = await call_error(
        "set_sheet_layout", path="sales.xlsx", sheet="Data", layout={"rows": actions}
    )
    assert "at most 10,000 rows" in message
    await call("set_sheet_layout", path="sales.xlsx", sheet="Data", layout={"rows": actions[:2]})
