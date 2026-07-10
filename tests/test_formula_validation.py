import os
import sys
import unittest

_REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
_SRC = os.path.join(_REPO_ROOT, "src")
if _SRC not in sys.path:
    sys.path.insert(0, _SRC)

from excel_mcp.validation import validate_formula  # noqa: E402


class TestValidateFormulaDenylist(unittest.TestCase):
    """Regression tests for the case-insensitive formula denylist.

    Excel function names are case-insensitive, so the unsafe-function
    denylist (INDIRECT, HYPERLINK, WEBSERVICE, DGET, RTD) must match
    mixed-case and lowercase forms as well as the canonical uppercase.
    """

    def test_uppercase_unsafe_blocked(self):
        for fn in ("INDIRECT", "HYPERLINK", "WEBSERVICE", "DGET", "RTD"):
            ok, msg = validate_formula(f"={fn}(A1)")
            self.assertFalse(ok, f"{fn} should be blocked: {msg}")
            self.assertIn(fn, msg)

    def test_lowercase_unsafe_blocked(self):
        for fn in ("indirect", "hyperlink", "webservice", "dget", "rtd"):
            ok, msg = validate_formula(f"={fn}(A1)")
            self.assertFalse(ok, f"{fn} (lowercase) should be blocked: {msg}")

    def test_mixed_case_unsafe_blocked(self):
        for fn in ("Indirect", "Hyperlink", "Webservice", "Dget", "Rtd"):
            ok, msg = validate_formula(f"={fn}(A1)")
            self.assertFalse(ok, f"{fn} (mixed case) should be blocked: {msg}")

    def test_safe_functions_allowed(self):
        for fn in ("SUM", "sum", "Sum", "AVERAGE", "average", "VLOOKUP", "vlookup"):
            ok, msg = validate_formula(f"={fn}(A1:A10)")
            self.assertTrue(ok, f"{fn} should be allowed: {msg}")

    def test_underscore_and_digit_function_names_blocked_when_unsafe(self):
        # Function names that start with a letter then contain
        # underscores/digits - confirms the broader pattern.
        ok, _ = validate_formula("=INDIRECT_X(A1)")
        self.assertTrue(ok)  # not in denylist
        ok, _ = validate_formula("=XLOOKUP(A1)")
        self.assertTrue(ok)  # not in denylist

    def test_no_functions_in_formula(self):
        ok, msg = validate_formula("=A1+B1")
        self.assertTrue(ok, f"plain arithmetic should be allowed: {msg}")


if __name__ == "__main__":
    unittest.main()
