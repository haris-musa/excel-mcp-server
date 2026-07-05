import os
import sys
import tempfile
import unittest

_REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
_SRC = os.path.join(_REPO_ROOT, "src")
if _SRC not in sys.path:
    sys.path.insert(0, _SRC)

import excel_mcp.server as server  # noqa: E402


class TestGetExcelPathSandbox(unittest.TestCase):
    def tearDown(self):
        server.EXCEL_FILES_PATH = None

    def test_stdio_accepts_absolute_only(self):
        server.EXCEL_FILES_PATH = None
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
            path = f.name
        try:
            self.assertEqual(server.get_excel_path(path), os.path.normpath(path))
            with self.assertRaises(ValueError):
                server.get_excel_path("relative_only.xlsx")
        finally:
            os.unlink(path)

    def test_remote_rejects_absolute(self):
        with tempfile.TemporaryDirectory() as d:
            server.EXCEL_FILES_PATH = d
            inner = os.path.join(d, "ok.xlsx")
            with self.assertRaises(ValueError):
                server.get_excel_path(inner)

    def test_remote_allows_relative_inside_sandbox(self):
        with tempfile.TemporaryDirectory() as d:
            server.EXCEL_FILES_PATH = d
            out = server.get_excel_path(os.path.join("subdir", "file.xlsx"))
            self.assertTrue(server._resolved_path_is_within(d, out))

    def test_remote_blocks_traversal(self):
        with tempfile.TemporaryDirectory() as d:
            server.EXCEL_FILES_PATH = d
            with self.assertRaises(ValueError):
                server.get_excel_path("../outside.xlsx")
            with self.assertRaises(ValueError):
                server.get_excel_path(os.path.join("a", "..", "..", "outside.xlsx"))

    def test_remote_rejects_nul(self):
        with tempfile.TemporaryDirectory() as d:
            server.EXCEL_FILES_PATH = d
            with self.assertRaises(ValueError):
                server.get_excel_path("a\x00b.xlsx")

    def test_stdio_rejects_shell_metacharacters(self):
        server.EXCEL_FILES_PATH = None
        bad_path = os.path.join(tempfile.gettempdir(), "safe.xlsx; echo injected")
        with self.assertRaises(ValueError):
            server.get_excel_path(bad_path)

    def test_remote_rejects_shell_metacharacters(self):
        with tempfile.TemporaryDirectory() as d:
            server.EXCEL_FILES_PATH = d
            for bad_path in [
                "safe.xlsx; echo injected",
                "safe.xlsx && touch injected",
                "safe.xlsx|cat",
                "`safe.xlsx`",
            ]:
                with self.subTest(bad_path=bad_path):
                    with self.assertRaises(ValueError):
                        server.get_excel_path(bad_path)


if __name__ == "__main__":
    unittest.main()
