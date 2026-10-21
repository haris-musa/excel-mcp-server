"""A sheet's XML, split into the sheet without its rows and the rows, each with its place."""

import re
from bisect import bisect_left, bisect_right
from dataclasses import dataclass, field
from functools import cached_property
from itertools import pairwise

from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.xml.constants import SHEET_MAIN_NS

from excel_mcp import lazy_shared

_SHEET_DATA = b"<sheetData"
_SHEET_DATA_END = b"</sheetData>"
_ROW = re.compile(rb"<row\b[^>]*>")
_ROW_NUMBER = re.compile(rb'\br="(\d+)"')
_HIDDEN = re.compile(rb'\bhidden="(?:1|true)"')
_CELL = re.compile(rb"<c[ >/]")
_CELL_COLUMN = re.compile(rb'<c\b[^>]*?\br="([A-Z]+)\d')
_UNADDRESSED = re.compile(rb"<row\b(?![^>]*\br=)|<c\b(?![^>]*\br=)")
_TAG = re.compile(rb"<(row|c)\b([^>]*)>")
_CELL_LETTERS = re.compile(rb'\br="([A-Z]+)\d')
_ROOT = re.compile(rb'<worksheet\b[^>]*?\bxmlns="([^"]+)"[^>]*>')


@dataclass
class SheetSource:
    """A sheet's XML and where each of its rows is in it."""

    data: bytes
    root: bytes
    row_numbers: list[int]
    spans: list[tuple[int, int]]
    hidden_rows: list[int]
    masters: dict[str, lazy_shared.SharedMaster] = field(default_factory=dict)

    def between(self, first: int, last: int) -> slice:
        """Where the rows numbered ``first`` to ``last`` are among the sheet's rows."""
        return slice(bisect_left(self.row_numbers, first), bisect_right(self.row_numbers, last))

    @cached_property
    def min_row(self) -> int:
        for number, (start, end) in zip(self.row_numbers, self.spans, strict=True):
            if _CELL.search(self.data, start, end):
                return number
        return 1

    @cached_property
    def max_row(self) -> int:
        for number, (start, end) in zip(
            reversed(self.row_numbers), reversed(self.spans), strict=True
        ):
            if _CELL.search(self.data, start, end):
                return number
        return 1

    @cached_property
    def columns(self) -> tuple[int, int]:
        """The first and last column that hold a cell."""
        indexes = {column_index_from_string(letters.decode()) for letters in self._letters}
        return (min(indexes), max(indexes)) if indexes else (1, 1)

    @property
    def max_column(self) -> int:
        return self.columns[1]

    @cached_property
    def _letters(self) -> set[bytes]:
        return set(_CELL_COLUMN.findall(self.data))

    @cached_property
    def has_cells(self) -> bool:
        return any(_CELL.search(self.data, start, end) for start, end in self.spans)


def split_sheet(data: bytes) -> tuple[bytes, SheetSource] | None:
    """The sheet without its rows, and the rows; None for a layout this does not handle.

    Rows and cells without a number or address are given one first, as their position
    says. The masters of shared formulas are indexed, as a row's formula can depend on the
    row that declares it.
    """
    root = _ROOT.search(data)
    start = data.find(_SHEET_DATA)
    end = data.find(_SHEET_DATA_END)
    if root is None or root[1] != SHEET_MAIN_NS.encode() or start < 0 or end < start:
        return None
    if not _all_addressed(data, start, end) and _UNADDRESSED.search(data, start, end):
        data = data[:start] + _address(data[start:end]) + data[end:]
        end = data.find(_SHEET_DATA_END)
    numbers: list[int] = []
    spans: list[tuple[int, int]] = []
    hidden: list[int] = []
    starts = list(_ROW.finditer(data, start, end))
    for position, tag in enumerate(starts):
        numbers.append(int(_ROW_NUMBER.search(tag.group())[1]))  # pyright: ignore[reportOptionalSubscript]
        following = starts[position + 1].start() if position + 1 < len(starts) else end
        spans.append((tag.start(), following))
        if _HIDDEN.search(tag.group()):
            hidden.append(numbers[-1])
    if numbers != sorted(set(numbers)):
        return None
    without_rows = data[:start] + _SHEET_DATA + b"/>" + data[end + len(_SHEET_DATA_END) :]
    masters = lazy_shared.index(data, start, end) if b't="shared"' in data else {}
    return without_rows, SheetSource(data, root.group(), numbers, spans, hidden, masters)


def _all_addressed(data: bytes, start: int, end: int) -> bool:
    """Whether every row and cell opens with its address, as counting them (quickly) shows."""
    return all(
        data.count(tag, start, end) == data.count(tag + b' r="', start, end)
        for tag in (b"<row", b"<c")
    )


def _address(rows: bytes) -> bytes:
    """The rows with the number of each row and the address of each cell that has none.

    Excel and openpyxl number them by position: a row after the one before it, a cell after
    the one before it in its row. Writers leave the numbers out to make the file smaller.
    """
    row = column = 0

    def addressed(tag: re.Match[bytes]) -> bytes:
        nonlocal row, column
        attributes = tag[2]
        if tag[1] == b"row":
            given = _ROW_NUMBER.search(attributes)
            row = int(given[1]) if given else row + 1
            column = 0
            return tag[0] if given else b'<row r="%d"%b>' % (row, attributes)
        given = _CELL_LETTERS.search(attributes)
        column = column_index_from_string(given[1].decode()) if given else column + 1
        if given:
            return tag[0]
        return b'<c r="%b%d"%b>' % (get_column_letter(column).encode(), row, attributes)

    bounds = [0, *(tag.start() for tag in _ROW.finditer(rows)), len(rows)]
    # One row at a time, so that no list of all the pieces of a sheet is held besides the sheet.
    return b"".join(_TAG.sub(addressed, rows[a:b]) for a, b in pairwise(bounds))
