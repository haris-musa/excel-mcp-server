from collections.abc import AsyncIterator, Callable, Coroutine, Iterator
from contextlib import contextmanager
from pathlib import Path
from typing import Any

import pytest
from mcp import Client
from mcp.types import CallToolResult, TextContent
from openpyxl import Workbook

from excel_mcp import lazy_edit, lazy_workbook
from excel_mcp.config import Settings
from excel_mcp.server import create_server

ToolCall = Callable[..., Coroutine[Any, Any, Any]]


@contextmanager
def lazily(monkeypatch: pytest.MonkeyPatch) -> Iterator[None]:
    """Every sheet is edited and read as a large one, in blocks of four rows."""
    with monkeypatch.context() as context:
        context.setattr(lazy_edit, "LAZY_SHEET_BYTES", 0)
        context.setattr(lazy_edit, "BLOCK_ROWS", 4)
        context.setattr(lazy_workbook, "BLOCK_ROWS", 4)
        yield


def pytest_addoption(parser: pytest.Parser) -> None:
    parser.addoption(
        "--lazy-sheets",
        action="store_true",
        help="Run every test with all sheets taken through the large-sheet path.",
    )


@pytest.fixture(autouse=True)
def forced_lazy_sheets(
    request: pytest.FixtureRequest, monkeypatch: pytest.MonkeyPatch
) -> Iterator[None]:
    if request.config.getoption("--lazy-sheets"):
        with lazily(monkeypatch):
            yield
    else:
        yield


def error_text(result: CallToolResult) -> str:
    content = result.content[0]
    assert isinstance(content, TextContent)
    return content.text


@pytest.fixture
def anyio_backend() -> str:
    return "asyncio"


@pytest.fixture
def files(tmp_path: Path) -> Path:
    directory = tmp_path / "files"
    directory.mkdir()
    return directory


@pytest.fixture
def sample(files: Path) -> Path:
    """A workbook with a small sales table on 'Data' and an empty 'Report' sheet."""
    workbook = Workbook()
    sheet = workbook.worksheets[0]
    sheet.title = "Data"
    for row in [
        ["Region", "Product", "Units", "Price"],
        ["North", "Apples", 10, 1.5],
        ["South", "Apples", 5, 1.5],
        ["North", "Pears", 7, 2.0],
        ["South", "Pears", 3, 2.0],
    ]:
        sheet.append(row)
    workbook.create_sheet("Report")
    path = files / "sales.xlsx"
    workbook.save(path)
    return path


@pytest.fixture
async def client(files: Path) -> AsyncIterator[Client]:
    server = create_server(Settings(allowed_dirs=[files]))
    async with Client(server) as connected:
        yield connected


@pytest.fixture
def call(client: Client) -> ToolCall:
    """Call a tool and return its structured result, failing the test on a tool error."""

    async def run(tool: str, /, **arguments: Any) -> Any:
        result = await client.call_tool(tool, arguments)
        assert not result.is_error, result.content
        structured = result.structured_content
        if structured is None:
            return error_text(result)
        if isinstance(structured, dict) and set(structured) == {"result"}:
            return structured["result"]
        return structured

    return run


@pytest.fixture
def call_error(client: Client) -> Callable[..., Coroutine[Any, Any, str]]:
    """Call a tool that must fail and return the error text."""

    async def run(tool: str, /, **arguments: Any) -> str:
        result = await client.call_tool(tool, arguments)
        assert result.is_error, result.structured_content
        return error_text(result)

    return run
