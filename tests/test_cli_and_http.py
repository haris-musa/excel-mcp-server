import os
from pathlib import Path
from typing import Any

import pytest

from excel_mcp import cli
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
