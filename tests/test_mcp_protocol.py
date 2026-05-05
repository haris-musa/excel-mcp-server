"""End-to-end MCP protocol smoke test.

Spawns the installed ``excel-mcp-server stdio`` binary as a real subprocess
and asserts that the JSON-RPC protocol works end-to-end:

* the MCP ``initialize`` handshake completes and reports a server name,
  protocol version, and ``tools`` capability; and
* a couple of documented tools (``create_workbook``, ``get_workbook_metadata``)
  are advertised in ``tools/list`` with valid object input schemas.

This complements the existing ``test_sandbox_paths.py`` (which exercises
internals directly) by verifying that the actual MCP entry point still
speaks valid JSON-RPC 2.0 — catching regressions in ``__main__.py``,
FastMCP API drift, or wiring issues that wouldn't fail any current test.

Skipped automatically if ``pytest-mcp-plugin`` is not installed or if
``excel-mcp-server`` is not on PATH, so this file is safe to land before
adding any new dev dependencies or workflows.

Install the optional dep with::

    uv add --dev "pytest-mcp-plugin>=0.2.3"

(0.2.3 sends ``notifications/initialized`` after the handshake, which
FastMCP requires before it will return tools from ``tools/list``.)
"""

from __future__ import annotations

import shutil
import unittest

try:
    from mcp_test import MCPTestClient
    HAS_MCP_TEST = True
except ImportError:
    HAS_MCP_TEST = False

HAS_BINARY = shutil.which("excel-mcp-server") is not None


@unittest.skipUnless(
    HAS_MCP_TEST,
    "install pytest-mcp-plugin to run: `uv add --dev pytest-mcp-plugin`",
)
@unittest.skipUnless(
    HAS_BINARY,
    "excel-mcp-server entry point not on PATH; run `uv sync` first",
)
class TestMCPProtocol(unittest.TestCase):
    def test_initialize_handshake_succeeds(self):
        """Server completes the MCP initialize handshake over stdio."""
        with MCPTestClient(
            ["excel-mcp-server", "stdio"],
            startup_timeout=15.0,
        ) as client:
            self.assertTrue(
                client.server_info.get("name"),
                "serverInfo.name must be set",
            )
            self.assertTrue(
                client.server_version,
                "server must advertise a protocolVersion",
            )
            self.assertIn(
                "tools",
                client.server_capabilities,
                "server must advertise the 'tools' capability",
            )

    def test_documented_tools_are_listed(self):
        """Tools documented in TOOLS.md are reachable via tools/list."""
        with MCPTestClient(
            ["excel-mcp-server", "stdio"],
            startup_timeout=15.0,
        ) as client:
            tools = client.list_tools()

            self.assertIsNotNone(
                tools.find("create_workbook"),
                "create_workbook tool must be advertised (documented in TOOLS.md)",
            )
            self.assertIsNotNone(
                tools.find("get_workbook_metadata"),
                "get_workbook_metadata tool must be advertised (documented in TOOLS.md)",
            )

            for tool in tools:
                self.assertTrue(
                    tool.input_schema,
                    f"{tool.name} missing inputSchema",
                )
                self.assertEqual(
                    tool.input_schema.get("type"),
                    "object",
                    f"{tool.name} inputSchema.type must be 'object'",
                )


if __name__ == "__main__":
    unittest.main()
