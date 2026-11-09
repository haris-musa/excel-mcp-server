"""A limit on the cell text one call may copy, expand, scan or return.

A cell limit alone does not bound a call: one cell holds up to 32,767 characters, and a
workbook can store one such text once and point every cell at it. So the work of a call that
builds something from cell text is counted in characters, as it is produced, and stops at the
first character over the limit.
"""

from excel_mcp.errors import LimitExceededError

CHARS_PER_CELL = 256
"""Text per cell that a limit on cells allows, on average; real data uses a tenth of it."""


class TextBudget:
    """The characters of text that may still be handled."""

    def __init__(self, cells: int, refusal: str) -> None:
        """``cells`` is the cell limit the call works under; ``refusal`` is the error message,
        in which ``{limit}`` stands for the limit in characters."""
        self.limit = cells * CHARS_PER_CELL
        self.left = self.limit
        self.refusal = refusal

    def spend(self, characters: int) -> None:
        self.left -= characters
        if self.left < 0:
            raise LimitExceededError(self.refusal.format(limit=f"{self.limit:,}"))

    def affords(self, characters: int) -> bool:
        """Spend ``characters`` if they fit; else spend nothing and say so."""
        if characters > self.left:
            return False
        self.left -= characters
        return True
