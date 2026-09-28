"""End-to-end runs of the real command over stdio and Streamable HTTP."""

import socket
import subprocess
import sys
import time
from pathlib import Path

import pytest
from mcp import Client, StdioServerParameters

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
