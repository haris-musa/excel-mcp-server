import os
from pathlib import Path
from typing import Any

import pytest

from excel_mcp import cli
from excel_mcp.config import Limits
from excel_mcp.server.http import BearerTokenMiddleware, is_loopback


def test_http_defaults_to_localhost_and_a_files_directory() -> None:
    args = cli._parser().parse_args(["streamable-http"])
    assert args.host == "127.0.0.1"
    assert cli._settings(args).allowed_dirs == [Path("excel_files")]


def test_stdio_is_unconfined_unless_directories_are_given(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    monkeypatch.delenv("EXCEL_FILES_PATH", raising=False)
    assert cli._settings(cli._parser().parse_args(["stdio"])).allowed_dirs == []

    monkeypatch.setenv("EXCEL_FILES_PATH", str(tmp_path) + os.pathsep)
    args = cli._parser().parse_args(["stdio", "--allow-dir", "other", "--read-only"])
    settings = cli._settings(args)
    assert settings.allowed_dirs == [Path("other"), tmp_path]
    assert settings.read_only


def test_limits_come_from_flags_and_environment(monkeypatch: pytest.MonkeyPatch) -> None:
    for name in [
        "EXCEL_MCP_MAX_FILE_MB",
        "EXCEL_MCP_MAX_CELLS",
        "EXCEL_MCP_MAX_PIVOT_CELLS",
        "EXCEL_MCP_MAX_LOADED_CELLS",
    ]:
        monkeypatch.delenv(name, raising=False)
    defaults = cli._settings(cli._parser().parse_args(["stdio"])).limits
    assert defaults == Limits()
    assert defaults.max_pivot_source_cells >= 2_000_000

    flags = ["stdio", "--max-file-mb", "7", "--max-cells", "11", "--max-pivot-cells", "13"]
    limits = cli._settings(cli._parser().parse_args([*flags, "--max-loaded-cells", "23"])).limits
    assert limits.max_loaded_cells == 23
    assert (limits.max_file_bytes, limits.max_cells, limits.max_pivot_source_cells) == (
        7 * 1024 * 1024,
        11,
        13,
    )

    monkeypatch.setenv("EXCEL_MCP_MAX_FILE_MB", "5")
    monkeypatch.setenv("EXCEL_MCP_MAX_CELLS", "17")
    monkeypatch.setenv("EXCEL_MCP_MAX_PIVOT_CELLS", "19")
    monkeypatch.setenv("EXCEL_MCP_MAX_LOADED_CELLS", "29")
    limits = cli._settings(cli._parser().parse_args(["stdio"])).limits
    assert limits.max_loaded_cells == 29
    assert (limits.max_file_bytes, limits.max_cells, limits.max_pivot_source_cells) == (
        5 * 1024 * 1024,
        17,
        19,
    )


@pytest.mark.parametrize("value", ["abc", "0", "-5"])
def test_bad_limit_values_are_usage_errors(
    value: str, monkeypatch: pytest.MonkeyPatch, capsys: pytest.CaptureFixture[str]
) -> None:
    monkeypatch.setenv("EXCEL_MCP_MAX_CELLS", value)
    with pytest.raises(SystemExit):
        cli._parser().parse_args(["stdio"])
    assert "argument --max-cells: expected a whole number above 0" in capsys.readouterr().err


def test_refuses_public_host_without_token(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.delenv("EXCEL_MCP_AUTH_TOKEN", raising=False)
    with pytest.raises(SystemExit, match="without authentication"):
        cli.main(["streamable-http", "--host", "0.0.0.0"])


@pytest.mark.parametrize(
    ("host", "expected"),
    [
        ("127.0.0.1", True),
        ("localhost", True),
        ("::1", True),
        ("0.0.0.0", False),
        ("127.0.0.2", False),
        ("10.0.0.5", False),
    ],
)
def test_is_loopback(host: str, expected: bool) -> None:
    assert is_loopback(host) is expected


async def _request(app: Any, headers: list[tuple[bytes, bytes]]) -> int:
    sent: list[dict[str, Any]] = []

    async def receive() -> dict[str, Any]:
        return {"type": "http.request", "body": b""}

    async def send(message: dict[str, Any]) -> None:
        sent.append(message)

    await app({"type": "http", "headers": headers}, receive, send)
    return sent[0]["status"]


@pytest.mark.anyio
async def test_bearer_token_middleware() -> None:
    async def ok_app(scope: Any, receive: Any, send: Any) -> None:
        await send({"type": "http.response.start", "status": 200, "headers": []})

    app = BearerTokenMiddleware(ok_app, "secret")
    assert await _request(app, []) == 401
    assert await _request(app, [(b"authorization", b"Bearer wrong")]) == 401
    assert await _request(app, [(b"authorization", b"Bearer secret")]) == 200
    assert await _request(app, [(b"authorization", b"bearer secret")]) == 200
