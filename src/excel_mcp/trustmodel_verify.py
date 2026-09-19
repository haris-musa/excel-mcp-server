"""Optional TrustModel AgentCert verify-gate for the HTTP transports.

Verifies a calling agent's **AgentCert + TrustScore** on each tool call, using the
dependency-free ``trustmodel-agentcert-tag`` client. It does **nothing** unless
``TRUSTMODEL_VERIFY=1`` is set; when enabled it defaults to **shadow mode** (logs a
verdict, never blocks). Set ``TRUSTMODEL_MODE=enforce`` to reject unverified/revoked
agents. Reads request metadata only (never tool payloads).

Docs: https://trustmodel.ai/verify
"""

from __future__ import annotations

import logging
import os

logger = logging.getLogger("excel-mcp")

_TRUTHY = {"1", "true", "yes", "on"}


def register_trustmodel_verify(mcp) -> bool:
    """Attach the verify-gate middleware when ``TRUSTMODEL_VERIFY`` is truthy.

    Safe to call unconditionally: a no-op when the flag is unset or the optional
    ``trustmodel-agentcert-tag`` dependency isn't installed. Returns ``True`` when
    the gate was attached.
    """
    if os.environ.get("TRUSTMODEL_VERIFY", "").strip().lower() not in _TRUTHY:
        return False

    try:
        from fastmcp.server.dependencies import get_http_headers
        from fastmcp.server.middleware import Middleware
        from trustmodel_agentcert_tag import extract_token, verify_gate
    except ImportError:
        logger.warning(
            "TRUSTMODEL_VERIFY is set but 'trustmodel-agentcert-tag' is not installed; "
            "install with `pip install excel-mcp-server[trustmodel]`. Skipping verify-gate."
        )
        return False

    mode = os.environ.get("TRUSTMODEL_MODE", "shadow")

    class _TrustModelVerifyMiddleware(Middleware):
        def __init__(self) -> None:
            self._guard = verify_gate(mode=mode)

        async def on_call_tool(self, context, call_next):
            # metadata-only: read the AgentCert token from the request headers.
            # shadow mode logs the verdict; enforce mode raises VerifyError.
            await self._guard(extract_token(get_http_headers() or {}))
            return await call_next(context)

    mcp.add_middleware(_TrustModelVerifyMiddleware())
    logger.info("TrustModel AgentCert verify-gate enabled (mode=%s).", mode)
    return True
