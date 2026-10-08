import random
import shutil
import zipfile
from collections.abc import AsyncIterator
from pathlib import Path

import pytest
from mcp import Client
from openpyxl import Workbook

from excel_mcp import cfb, ovba
from excel_mcp.config import Settings
from excel_mcp.ovba_compress import compress
from excel_mcp.ovba_write import Project
from excel_mcp.server import create_server
from tests.conftest import ToolCall, error_text

pytestmark = pytest.mark.anyio

FIXTURES = Path(__file__).parent / "fixtures"
MACRO = 'Sub Hello()\n    Range("A1").Value = "hi"\nEnd Sub'


@pytest.fixture
async def vba_client(files: Path) -> AsyncIterator[Client]:
    server = create_server(Settings(allowed_dirs=[files], allow_vba_write=True))
    async with Client(server) as connected:
        yield connected


@pytest.fixture
def vba_call(vba_client: Client) -> ToolCall:
    async def run(tool: str, /, **arguments: object) -> object:
        result = await vba_client.call_tool(tool, arguments)
        assert not result.is_error, result.content
        return result.structured_content

    return run


@pytest.fixture
def vba_error(vba_client: Client) -> ToolCall:
    async def run(tool: str, /, **arguments: object) -> str:
        result = await vba_client.call_tool(tool, arguments)
        assert result.is_error, result.structured_content
        return error_text(result)

    return run


@pytest.fixture
def macro_workbook(files: Path) -> Path:
    return Path(shutil.copy(FIXTURES / "macro01.xlsm", files / "macros.xlsm"))


async def _modules(call: ToolCall, path: str) -> dict[str, dict]:
    project = await call("read_vba", path=path)
    return {module["name"]: module for module in project["modules"]}


async def test_tools_are_absent_by_default_and_in_read_only_mode(files: Path) -> None:
    names = {"write_vba_module", "delete_vba_module"}
    for settings in (
        Settings(allowed_dirs=[files]),
        Settings(allowed_dirs=[files], read_only=True, allow_vba_write=True),
    ):
        async with Client(create_server(settings)) as client:
            assert not names & {tool.name for tool in (await client.list_tools()).tools}
    async with Client(
        create_server(Settings(allowed_dirs=[files], allow_vba_write=True))
    ) as client:
        tools = {tool.name: tool for tool in (await client.list_tools()).tools}
    assert names <= set(tools)
    for name in names:
        hints = tools[name].annotations
        assert hints and hints.destructive_hint and not hints.read_only_hint


async def test_macro_workbooks_cannot_be_created_unless_vba_writing_is_allowed(
    client: Client, files: Path
) -> None:
    result = await client.call_tool("create_workbook", {"path": "m.xlsm"})
    assert result.is_error
    assert "cannot contain macros" in error_text(result)
    assert not (files / "m.xlsm").exists()


async def test_created_macro_workbook_gets_a_project_with_document_modules(
    vba_call: ToolCall, call: ToolCall, files: Path
) -> None:
    await vba_call("create_workbook", path="m.xlsm", sheets=["Data", "Report"])
    modules = await _modules(call, "m.xlsm")
    assert {name: module["kind"] for name, module in modules.items()} == {
        "ThisWorkbook": "document",
        "Sheet1": "document",
        "Sheet2": "document",
    }
    with zipfile.ZipFile(files / "m.xlsm") as archive:
        assert "xl/vbaProject.bin" in archive.namelist()
        assert "macroEnabled" in archive.read("[Content_Types].xml").decode()
        assert 'codeName="Sheet2"' in archive.read("xl/worksheets/sheet2.xml").decode()


async def test_write_creates_a_standard_module_and_reads_back_exactly(
    vba_call: ToolCall, call: ToolCall
) -> None:
    await vba_call("create_workbook", path="m.xlsm")
    change = await vba_call("write_vba_module", path="m.xlsm", module="Module1", code=MACRO)
    assert change == {"module": "Module1", "kind": "standard", "action": "created"}
    modules = await _modules(call, "m.xlsm")
    assert modules["Module1"]["code"] == MACRO
    assert modules["Module1"]["kind"] == "standard"
    assert modules["Module1"]["line_count"] == 3


async def test_write_replaces_a_module_and_keeps_the_others(
    vba_call: ToolCall, call: ToolCall, macro_workbook: Path
) -> None:
    before = await _modules(call, "macros.xlsm")
    change = await vba_call("write_vba_module", path="macros.xlsm", module="module1", code=MACRO)
    assert change == {"module": "Module1", "kind": "standard", "action": "replaced"}
    after = await _modules(call, "macros.xlsm")
    assert after["Module1"]["code"] == MACRO
    assert list(after) == list(before)
    for name in before:
        if name != "Module1":
            assert after[name] == before[name]


async def test_document_module_code_can_be_set_and_keeps_its_kind(
    vba_call: ToolCall, call: ToolCall
) -> None:
    await vba_call("create_workbook", path="m.xlsm")
    handler = 'Private Sub Workbook_Open()\n    Sheet1.Range("A1").Value = 1\nEnd Sub'
    change = await vba_call("write_vba_module", path="m.xlsm", module="ThisWorkbook", code=handler)
    assert change == {"module": "ThisWorkbook", "kind": "document", "action": "replaced"}
    modules = await _modules(call, "m.xlsm")
    assert modules["ThisWorkbook"]["code"] == handler
    assert modules["ThisWorkbook"]["kind"] == "document"


async def test_class_module_and_delete(vba_call: ToolCall, call: ToolCall) -> None:
    await vba_call("create_workbook", path="m.xlsm")
    await vba_call("write_vba_module", path="m.xlsm", module="Module1", code=MACRO)
    created = await vba_call(
        "write_vba_module", path="m.xlsm", module="Counter", code="Public N As Long", kind="class"
    )
    assert created["kind"] == "class"
    assert (await _modules(call, "m.xlsm"))["Counter"]["kind"] == "class"

    deleted = await vba_call("delete_vba_module", path="m.xlsm", module="counter")
    assert deleted == {"module": "Counter", "kind": "class", "action": "deleted"}
    await vba_call("delete_vba_module", path="m.xlsm", module="Module1")
    assert list(await _modules(call, "m.xlsm")) == ["ThisWorkbook", "Sheet1"]


async def test_delete_rejects_document_modules_and_unknown_names(
    vba_error: ToolCall, vba_call: ToolCall
) -> None:
    await vba_call("create_workbook", path="m.xlsm")
    assert "Replace its code" in await vba_error(
        "delete_vba_module", path="m.xlsm", module="ThisWorkbook"
    )
    assert "Modules: 'ThisWorkbook'" in await vba_error(
        "delete_vba_module", path="m.xlsm", module="Nope"
    )


async def test_xlsx_is_refused(vba_error: ToolCall, sample: Path) -> None:
    message = await vba_error("write_vba_module", path="sales.xlsx", module="M", code=MACRO)
    assert "only be stored in .xlsm or .xltm" in message


@pytest.mark.parametrize(
    ("module", "code", "expected"),
    [
        ("Module1", 'Attribute VB_Name = "x"\nSub A()\nEnd Sub', "cannot start with 'Attribute'"),
        ("Module1", 'Sub A()\rAttribute VB_Name = "x"\rEnd Sub', "cannot start with 'Attribute'"),
        (
            "Module1",
            'Sub A()\r\nAttribute VB_Name = "x"\r\nEnd Sub',
            "cannot start with 'Attribute'",
        ),
        ("Module1", 'Sub A()\n  attribute VB_Name = "x"\nEnd Sub', "cannot start with 'Attribute'"),
        ("Module1", "Sub A()\0End Sub", "NUL"),
        ("Module1", "x = 1 中", "ChrW"),
        ("1bad", MACRO, "not a valid module name"),
        ("A" * 32, MACRO, "not a valid module name"),
        ("my module", MACRO, "not a valid module name"),
    ],
)
async def test_invalid_code_or_names_are_rejected(
    vba_error: ToolCall, vba_call: ToolCall, module: str, code: str, expected: str
) -> None:
    await vba_call("create_workbook", path="m.xlsm")
    assert expected in await vba_error("write_vba_module", path="m.xlsm", module=module, code=code)


async def test_oversized_code_is_rejected(vba_client: Client, vba_call: ToolCall) -> None:
    await vba_call("create_workbook", path="m.xlsm")
    result = await vba_client.call_tool(
        "write_vba_module", {"path": "m.xlsm", "module": "M", "code": "x" * 200_001}
    )
    assert result.is_error


async def test_paths_outside_the_allowed_directory_are_rejected(vba_error: ToolCall) -> None:
    for path, expected in (("../escape.xlsm", "outside"), ("notes.txt", "Excel file")):
        message = await vba_error("write_vba_module", path=path, module="M", code=MACRO)
        assert expected in message
    assert "outside" in await vba_error("delete_vba_module", path="../escape.xlsm", module="M")


async def test_later_edits_keep_the_macros(vba_call: ToolCall, call: ToolCall) -> None:
    await vba_call("create_workbook", path="m.xlsm")
    await vba_call("write_vba_module", path="m.xlsm", module="Module1", code=MACRO)
    await vba_call("write_range", path="m.xlsm", sheet="Sheet1", at="A1", rows=[[1, 2]])
    assert (await _modules(call, "m.xlsm"))["Module1"]["code"] == MACRO
    await vba_call("write_vba_module", path="m.xlsm", module="Module2", code="Public X As Long")
    assert (await call("read_range", path="m.xlsm", sheet="Sheet1", range="A1:B1"))["values"] == [
        [1, 2]
    ]


async def test_macro_workbook_without_a_project_gets_one(
    vba_call: ToolCall, call: ToolCall, files: Path
) -> None:
    workbook = Workbook()
    workbook.active.title = "Data"  # pyright: ignore[reportOptionalMemberAccess]
    workbook.save(files / "bare.xlsm")
    await vba_call("write_vba_module", path="bare.xlsm", module="Module1", code=MACRO)
    modules = await _modules(call, "bare.xlsm")
    assert list(modules) == ["ThisWorkbook", "Sheet1", "Module1"]
    with zipfile.ZipFile(files / "bare.xlsm") as archive:
        assert "macroEnabled" in archive.read("[Content_Types].xml").decode()
        assert "vbaProject" in archive.read("xl/_rels/workbook.xml.rels").decode()


async def test_large_code_spans_several_compression_chunks(
    vba_call: ToolCall, call: ToolCall
) -> None:
    rng = random.Random(1)
    lines = [f"' {rng.getrandbits(64):x} {rng.getrandbits(64):x}" for _ in range(3000)]
    code = "Sub Big()\n" + "\n".join(lines) + "\nEnd Sub"
    await vba_call("create_workbook", path="m.xlsm")
    await vba_call("write_vba_module", path="m.xlsm", module="Big", code=code)
    project = await call("read_vba", path="m.xlsm", max_chars=200_000)
    assert {m["name"]: m for m in project["modules"]}["Big"]["code"] == code


def test_rewriting_a_project_keeps_modules_and_drops_compiled_code() -> None:
    original = (FIXTURES / "vbaProject.bin").read_bytes()
    before = ovba.read_project(original)
    rewritten = Project.from_bytes(original).to_bytes()
    after = ovba.read_project(rewritten)
    assert [(m.name, m.kind, m.source) for m in after.modules] == [
        (m.name, m.kind, m.source) for m in before.modules
    ]
    assert after.text == before.text
    assert after.header == before.header
    paths = {entry.path for entry in cfb.read_entries(rewritten)}
    assert not any("__SRP_" in path for path in paths)
    assert {"PROJECT", "PROJECTwm", "VBA/dir", "VBA/_VBA_PROJECT"} <= paths


def test_compound_files_round_trip() -> None:
    entries = [
        cfb.Entry("VBA"),
        cfb.Entry("VBA/dir", b"d" * 10),
        cfb.Entry("VBA/big", bytes(range(256)) * 40),
        cfb.Entry("Form", clsid="C62A69F0-16DC-11CE-9E98-00AA00574A4F"),
        cfb.Entry("Form/f/o", b"z" * 5000),
        cfb.Entry("empty", b""),
        *(cfb.Entry(f"s{n}", bytes([n]) * n) for n in range(1, 12)),
    ]
    content = cfb.write(entries)
    read_back = {entry.path: entry for entry in cfb.read_entries(content)}
    for entry in entries:
        assert read_back[entry.path].data == entry.data
    assert read_back["Form"].data is None
    assert read_back["Form"].clsid.upper() == "C62A69F0-16DC-11CE-9E98-00AA00574A4F"


@pytest.mark.parametrize(
    "data",
    [
        b"",
        b"a",
        b"abcabcabcabcabcabc" * 500,
        b"\r\n".join(b"Sub X%d()" % n for n in range(2000)),
        random.Random(7).randbytes(9000),
    ],
)
def test_compression_round_trips(data: bytes) -> None:
    assert ovba.decompress(compress(data)).rstrip(b"\0") == data.rstrip(b"\0")
    if len(data) % 4096 == 0:
        assert ovba.decompress(compress(data)) == data


def test_compression_makes_repetitive_code_smaller() -> None:
    data = b'Range("A1").Value = 1\r\n' * 400
    assert len(compress(data)) < len(data) // 4


def test_class_modules_get_the_attributes_excel_needs() -> None:
    # Without VB_Base Excel cannot load a class module that is not Excel's own.
    project = Project.new([("ThisWorkbook", "{00020819-0000-0000-C000-000000000046}")])
    project.set_code("Counter", "Public Total As Long", "class")
    module = project.find("Counter")
    assert module is not None
    assert 'Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"' in module.source
    assert "Class=Counter" in project.text


def test_vba_write_flag_comes_from_the_command_line_or_environment(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from excel_mcp import cli

    monkeypatch.delenv("EXCEL_MCP_ALLOW_VBA_WRITE", raising=False)
    assert not cli._settings(cli._parser().parse_args(["stdio"])).allow_vba_write
    assert cli._settings(cli._parser().parse_args(["stdio", "--allow-vba-write"])).allow_vba_write
    monkeypatch.setenv("EXCEL_MCP_ALLOW_VBA_WRITE", "1")
    assert cli._settings(cli._parser().parse_args(["stdio"])).allow_vba_write


@pytest.mark.parametrize("tool", ["write_vba_module", "delete_vba_module"])
async def test_signed_projects_are_refused_and_left_untouched(
    vba_error: ToolCall, macro_workbook: Path, tool: str
) -> None:
    with zipfile.ZipFile(macro_workbook, "a") as archive:
        archive.writestr("xl/vbaProjectSignature.bin", b"signature")
    before = macro_workbook.read_bytes()
    arguments = {"code": MACRO} if tool == "write_vba_module" else {}
    message = await vba_error(tool, path="macros.xlsm", module="Module1", **arguments)
    assert "digitally signed" in message
    assert macro_workbook.read_bytes() == before
