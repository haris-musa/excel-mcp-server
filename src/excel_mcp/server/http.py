"""Serving over Streamable HTTP, optionally behind a bearer token."""

import hmac

import uvicorn
from mcp.server import MCPServer
from starlette.types import ASGIApp, Receive, Scope, Send


class BearerTokenMiddleware:
    """Reject HTTP requests that do not carry ``Authorization: Bearer <token>``."""

    def __init__(self, app: ASGIApp, token: str) -> None:
        self.app = app
        self.token = token.encode()

    async def __call__(self, scope: Scope, receive: Receive, send: Send) -> None:
        if scope["type"] == "http":
            header = dict(scope["headers"]).get(b"authorization", b"")
            scheme, _, credentials = header.partition(b" ")
            if scheme.lower() != b"bearer" or not hmac.compare_digest(credentials, self.token):
                await _unauthorized(send)
                return
        await self.app(scope, receive, send)


async def _unauthorized(send: Send) -> None:
    await send(
        {
            "type": "http.response.start",
            "status": 401,
            "headers": [
                (b"content-type", b"application/json"),
                (b"www-authenticate", b"Bearer"),
            ],
        }
    )
    await send({"type": "http.response.body", "body": b'{"error":"unauthorized"}'})


# The hosts for which the MCP SDK enables DNS rebinding protection. Other
# loopback addresses such as 127.0.0.2 would be reachable through a rebound
# domain without it, so they need a token like any public address.
LOOPBACK_HOSTS = frozenset({"127.0.0.1", "localhost", "::1"})


def is_loopback(host: str) -> bool:
    return host in LOOPBACK_HOSTS


def build_app(server: MCPServer, host: str, token: str | None) -> ASGIApp:
    app = server.streamable_http_app(host=host)
    return BearerTokenMiddleware(app, token) if token else app


def serve(server: MCPServer, host: str, port: int, token: str | None) -> None:
    uvicorn.run(build_app(server, host, token), host=host, port=port, log_level="warning")
