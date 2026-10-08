"""Rows or columns inserted into or deleted from a sheet, and where a position goes."""

from dataclasses import dataclass
from typing import Literal

from excel_mcp.refs import MAX_COLUMN, MAX_ROW, CellRange

Axis = Literal["rows", "columns"]


@dataclass(frozen=True)
class LineEdit:
    """``count`` rows or columns inserted before line ``at``, or deleted from ``at`` on.

    Lines are 1-based. ``sheet`` is the title of the sheet that changes.
    """

    sheet: str
    axis: Axis
    at: int
    count: int
    delete: bool

    @property
    def limit(self) -> int:
        return MAX_ROW if self.axis == "rows" else MAX_COLUMN

    @property
    def end(self) -> int:
        """The last deleted line."""
        return self.at + self.count - 1

    def index(self, line: int) -> int | None:
        """Where a line goes, or None when it is deleted."""
        if not self.delete:
            return line + self.count if line >= self.at else line
        if line < self.at:
            return line
        return None if line <= self.end else line - self.count

    def corner(self, line: int) -> tuple[int, bool]:
        """Where the edge of an object that sits in ``line`` goes, and whether the offset
        into the line is lost. Excel puts an edge whose line is deleted at the top of the
        line that now takes the deleted ones' place."""
        moved = self.index(line)
        return (self.at, True) if moved is None else (moved, False)

    def span(self, first: int, last: int) -> tuple[int, int] | None:
        """Where lines ``first`` to ``last`` end up, or None when all are deleted.

        A first line that is deleted gives way to the next one, a last line to the previous.
        """
        if self.delete and self.at <= first and last <= self.end:
            return None
        if not self.delete:
            return self._pushed(first), min(self._pushed(last), self.limit)
        return self._after(first, self.at), self._after(last, self.at - 1)

    def start(self, line: int) -> int:
        """Where the first line of a span goes; a deleted one gives way to the next."""
        return self._pushed(line) if not self.delete else self._after(line, self.at)

    def stop(self, line: int) -> int:
        """Where the last line of a span goes; a deleted one gives way to the previous."""
        if not self.delete:
            return min(self._pushed(line), self.limit)
        return self._after(line, self.at - 1)

    def cuts(self, first: int, last: int) -> bool:
        """Whether the edit would split the lines ``first`` to ``last`` into separate parts."""
        if not self.delete:
            return first < self.at <= last
        overlaps = first <= self.end and self.at <= last
        return overlaps and not (self.at <= first and last <= self.end)

    def range(self, area: CellRange) -> CellRange | None:
        """Where a block of cells ends up; it does not grow over inserted lines."""
        return self.area(area, grow=False)

    def source(self, line: int) -> int:
        """The line of the old sheet that a line of the new one is, or takes its size from.

        Inserted lines are formatted like the line above them.
        """
        if self.delete:
            return line if line < self.at else line + self.count
        if line < self.at:
            return line
        return line - self.count if line >= self.at + self.count else max(self.at - 1, 1)

    def area(self, area: CellRange, *, grow: bool = True) -> CellRange | None:
        """Where a block of cells ends up. A block that ends just above an insertion grows
        over the inserted lines, as ranges of formats and validation do in Excel, unless
        ``grow`` is False (for a single cell, which only moves)."""
        rows = self.axis == "rows"
        first, last = (area.min_row, area.max_row) if rows else (area.min_col, area.max_col)
        moved = self.span(first, last)
        if moved is None:
            return None
        if grow and not self.delete and last == self.at - 1:
            moved = (moved[0], moved[1] + self.count)
        if rows:
            return CellRange(moved[0], area.min_col, moved[1], area.max_col)
        return CellRange(area.min_row, moved[0], area.max_row, moved[1])

    def _pushed(self, line: int) -> int:
        return line + self.count if line >= self.at else line

    def _after(self, line: int, inside: int) -> int:
        """``inside`` is where an edge goes when it lay in the deleted lines."""
        if line < self.at:
            return line
        if line <= self.end:
            return inside
        return line if line == self.limit else line - self.count
