"""Macro workbooks are stored only when the server may write VBA."""

import base64
from pathlib import Path

import pytest
from mcp import Client

from excel_mcp.config import Settings
from excel_mcp.server import create_server
from tests.conftest import ToolCall
from tests.package_support import read_parts

pytestmark = pytest.mark.anyio

FIXTURES = Path(__file__).parent / "fixtures"


def _encoded(path: Path) -> str:
    return base64.b64encode(path.read_bytes()).decode()


@pytest.mark.parametrize("name", ["evil.xlsm", "evil.xltm", "EVIL.XLSM"])
async def test_macro_workbooks_cannot_be_imported_without_vba_writing(
    call_error: ToolCall, files: Path, name: str
) -> None:
    content = _encoded(FIXTURES / "macro01.xlsm")
    message = await call_error("import_workbook", path=name, content_base64=content)
    assert "--allow-vba-write" in message
    assert not list(files.iterdir())


async def test_macro_free_uploads_follow_the_same_rule_as_create(
    call_error: ToolCall, sample: Path
) -> None:
    content = _encoded(sample)
    message = await call_error(
        "import_workbook", path="plain.xlsm", content_base64=content, overwrite=True
    )
    assert "--allow-vba-write" in message


async def test_macros_are_dropped_from_uploads_to_plain_workbooks(
    call: ToolCall, files: Path
) -> None:
    content = _encoded(FIXTURES / "macro01.xlsm")
    await call("import_workbook", path="plain.xlsx", content_base64=content)
    assert not [name for name in read_parts(files / "plain.xlsx") if "vbaProject" in name]


async def test_macro_uploads_are_stored_when_vba_writing_is_allowed(files: Path) -> None:
    settings = Settings(allowed_dirs=[files], allow_vba_write=True)
    async with Client(create_server(settings)) as client:
        result = await client.call_tool(
            "import_workbook",
            {"path": "ok.xlsm", "content_base64": _encoded(FIXTURES / "macro01.xlsm")},
        )
    assert not result.is_error
    assert "xl/vbaProject.bin" in read_parts(files / "ok.xlsm")
