"""Edits that would grow a workbook past what the server accepts are refused."""

from pathlib import Path

import pytest
from mcp import Client
from openpyxl import load_workbook
from PIL import Image

from excel_mcp.config import Limits, Settings
from excel_mcp.server import create_server
from tests.conftest import error_text

pytestmark = pytest.mark.anyio


async def _error(files: Path, limits: Limits, tool: str, **arguments: object) -> str:
    async with Client(create_server(Settings(allowed_dirs=[files], limits=limits))) as client:
        result = await client.call_tool(tool, arguments)
    assert result.is_error
    return error_text(result)


async def test_copy_sheet_refuses_sheets_with_too_many_cells(sample: Path, files: Path) -> None:
    message = await _error(
        files,
        Limits(max_copy_cells=10),
        "copy_sheet",
        path="sales.xlsx",
        sheet="Data",
        new_name="C",
    )
    assert "20 cells; at most 10" in message
    assert load_workbook(sample).sheetnames == ["Data", "Report"]


async def test_edits_that_make_the_file_too_large_leave_it_unchanged(
    sample: Path, files: Path
) -> None:
    Image.effect_noise((300, 300), 90).convert("RGB").save(files / "noise.png")
    before = sample.read_bytes()
    message = await _error(
        files,
        Limits(max_file_bytes=len(before) + 1000),
        "insert_image",
        path="sales.xlsx",
        sheet="Report",
        image_path="noise.png",
        at="A1",
    )
    assert "limit is" in message
    assert sample.read_bytes() == before
