"""Regression test for isError-compliance in excel-mcp-server tools.

The bug class: tool handlers caught their domain exceptions (ValidationError,
CalculationError, etc.) and returned a formatted 'Error: ...' string, while
letting unexpected exceptions fall through to a second 'except Exception' that
re-raised. FastMCP treats the return value as success content with
isError=false, so MCP clients (Claude Code, Cursor, etc.) and LLMs see the
error text as data, not as a tool failure.

The fix replaces `return f\"Error: {str(e)}\"` with bare `raise`, which
preserves the original exception type so FastMCP sets isError=true on the
wire response while still surfacing the message in the content text.

Reference: https://composio.dev/blog/mcp-security-vulnerabilities (Dayna
Blackwell's MCP security audit, June 2026).
"""
import asyncio
import os
import sys
import tempfile
import unittest

_REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
_SRC = os.path.join(_REPO_ROOT, "src")
if _SRC not in sys.path:
    sys.path.insert(0, _SRC)

from mcp.server.fastmcp.exceptions import ToolError

import excel_mcp.server as server  # noqa: E402


class TestIsErrorCompliance(unittest.TestCase):
    """Tool failures must surface as isError=true, not as success content."""

    def test_validate_formula_syntax_on_missing_file_raises(self) -> None:
        """A ValidationError on a missing Excel file must raise ToolError,
        not return a formatted 'Error: ...' string."""
        server.EXCEL_FILES_PATH = None  # use stdio mode (absolute paths)
        with tempfile.TemporaryDirectory() as d:
            nonexistent = os.path.join(d, "does_not_exist.xlsx")
            with self.assertRaises(ToolError) as ctx:
                asyncio.run(
                    server.mcp.call_tool(
                        "validate_formula_syntax",
                        {
                            "filepath": nonexistent,
                            "sheet_name": "Sheet1",
                            "cell": "A1",
                            "formula": "=1+1",
                        },
                    )
                )
            # The ToolError message should mention the failure mode.
            self.assertIn("Error", str(ctx.exception))

    def test_create_workbook_in_existing_dir_still_succeeds(self) -> None:
        """A successful tool call must return content (not raise) and
        NOT carry isError=true on the wire."""
        server.EXCEL_FILES_PATH = None
        with tempfile.TemporaryDirectory() as d:
            target = os.path.join(d, "new.xlsx")
            content, _ = asyncio.run(
                server.mcp.call_tool("create_workbook", {"filepath": target})
            )
            self.assertTrue(content, "expected non-empty content list")
            self.assertIn("Created workbook", content[0].text)
            self.assertTrue(os.path.exists(target))
            os.unlink(target)


if __name__ == "__main__":
    unittest.main()
