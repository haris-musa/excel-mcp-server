import contextlib
import random
import shutil
import zipfile
from pathlib import Path

import pytest

from excel_mcp import ovba
from excel_mcp.errors import WorkbookError
from tests.conftest import ToolCall

pytestmark = pytest.mark.anyio

FIXTURES = Path(__file__).parent / "fixtures"


@pytest.fixture
def macro_workbook(files: Path) -> Path:
    return Path(shutil.copy(FIXTURES / "macro01.xlsm", files / "macros.xlsm"))


async def test_read_vba_lists_modules_with_code(call: ToolCall, macro_workbook: Path) -> None:
    project = await call("read_vba", path="macros.xlsm")
    assert [(module["name"], module["kind"]) for module in project["modules"]] == [
        ("Module1", "standard"),
        ("ThisWorkbook", "document"),
        ("Sheet1", "document"),
    ]
    module1 = project["modules"][0]
    assert module1["code"].startswith("Sub ")
    assert "Attribute VB_" not in module1["code"]
    assert module1["line_count"] == len(module1["code"].splitlines())
    assert project["truncated"] is False


async def test_read_vba_single_module(
    call: ToolCall, call_error: ToolCall, macro_workbook: Path
) -> None:
    project = await call("read_vba", path="macros.xlsm", module="module1")
    assert [module["name"] for module in project["modules"]] == ["Module1"]
    message = await call_error("read_vba", path="macros.xlsm", module="Nope")
    assert "'Module1', 'ThisWorkbook', 'Sheet1'" in message


async def test_read_vba_respects_the_character_budget(call: ToolCall, macro_workbook: Path) -> None:
    project = await call("read_vba", path="macros.xlsm", max_chars=5)
    assert project["truncated"] is True
    assert sum(len(module["code"]) for module in project["modules"]) == 5


async def test_read_vba_on_a_workbook_without_macros(call_error: ToolCall, sample: Path) -> None:
    assert "no VBA macros" in await call_error("read_vba", path="sales.xlsx")


async def test_describe_workbook_reports_macros(
    call: ToolCall, sample: Path, macro_workbook: Path
) -> None:
    assert (await call("describe_workbook", path="macros.xlsm"))["has_vba"] is True
    assert (await call("describe_workbook", path="sales.xlsx"))["has_vba"] is False


def test_excel_made_project_is_decoded() -> None:
    modules, project_text = ovba.read_project((FIXTURES / "vbaProject.bin").read_bytes())
    module1 = next(module for module in modules if module.name == "Module1")
    assert module1.procedural
    assert 'MsgBox ("Hello from Python!")' in module1.source
    assert "Module=Module1" in project_text


def test_decompress_uncompressed_chunk() -> None:
    raw = bytes(range(256)) * 16
    header = (0x3FFF).to_bytes(2, "little")
    assert ovba.decompress(b"\x01" + header + raw) == raw


@pytest.mark.parametrize(
    "data",
    [
        b"",
        b"\x02\x00\x00",
        b"\x01\x05",
        b"\x01\x04\xb0\x02\x00\x10",
    ],
)
def test_decompress_rejects_invalid_data(data: bytes) -> None:
    with pytest.raises(WorkbookError):
        ovba.decompress(data)


def test_non_ole_project_is_rejected() -> None:
    with pytest.raises(WorkbookError, match="OLE"):
        ovba.read_project(b"not an OLE file")


async def test_damaged_project_gives_a_clear_error(
    call_error: ToolCall, files: Path, macro_workbook: Path
) -> None:
    damaged = files / "damaged.xlsm"
    with zipfile.ZipFile(macro_workbook) as source, zipfile.ZipFile(damaged, "w") as target:
        for item in source.infolist():
            content = source.read(item)
            if item.filename == "xl/vbaProject.bin":
                content = content[:1024]
            target.writestr(item, content)
    message = await call_error("read_vba", path="damaged.xlsm")
    assert "The VBA project" in message


def test_corrupted_projects_only_raise_workbook_errors() -> None:
    original = (FIXTURES / "vbaProject.bin").read_bytes()
    rng = random.Random(1)
    for _ in range(300):
        mutated = bytearray(original)
        for _ in range(rng.randint(1, 8)):
            mutated[rng.randrange(len(mutated))] = rng.randrange(256)
        with contextlib.suppress(WorkbookError):
            ovba.read_project(bytes(mutated))


def test_oversized_sector_header_is_rejected_before_parsing() -> None:
    header = bytearray((FIXTURES / "vbaProject.bin").read_bytes())
    header[30:32] = (0xFFFF).to_bytes(2, "little")
    with pytest.raises(WorkbookError, match="OLE"):
        ovba.read_project(bytes(header))
