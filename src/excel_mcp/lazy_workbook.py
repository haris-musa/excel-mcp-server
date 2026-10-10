"""Workbooks whose cells are parsed, row by row, when something asks for them.

Calculating a page of a large sheet needs the cells of that page and of what its formulas
use, not the other million. The workbook is loaded without its cell data, and each sheet
keeps its XML so rows can be parsed on demand with openpyxl's own cell parser. Only the
cells and the hidden state of rows are loaded this way: that is all the calculator reads.
"""

import re
from bisect import bisect_left, bisect_right
from collections.abc import Iterator
from dataclasses import dataclass
from functools import cached_property
from io import BytesIO
from pathlib import Path
from typing import IO, Any
from zipfile import ZipFile

from openpyxl import Workbook
from openpyxl.cell import Cell, MergedCell
from openpyxl.reader.excel import ExcelReader
from openpyxl.utils import column_index_from_string
from openpyxl.worksheet._reader import WorkSheetParser
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.xml.constants import WORKSHEET_TYPE

BLOCK_ROWS = 128
_SHEET_DATA = b"<sheetData"
_SHEET_DATA_END = b"</sheetData>"
_ROW = re.compile(rb"<row\b[^>]*>")
_ROW_NUMBER = re.compile(rb'\br="(\d+)"')
_HIDDEN = re.compile(rb'\bhidden="(?:1|true)"')
_CELL = re.compile(rb"<c[ >/]")
_CELL_COLUMN = re.compile(rb'<c\b[^>]*?\br="([A-Z]+)\d')
_NAMESPACE = re.compile(rb'<worksheet\b[^>]*?\bxmlns="([^"]+)"')

CellKey = tuple[int, int]


@dataclass
class SheetSource:
    """A sheet's XML and where each of its rows is in it."""

    data: bytes
    namespace: bytes
    row_numbers: list[int]
    spans: list[tuple[int, int]]
    hidden_rows: list[int]

    def between(self, first: int, last: int) -> slice:
        """Where the rows numbered ``first`` to ``last`` are among the sheet's rows."""
        return slice(bisect_left(self.row_numbers, first), bisect_right(self.row_numbers, last))

    @cached_property
    def max_row(self) -> int:
        for number, (start, end) in zip(
            reversed(self.row_numbers), reversed(self.spans), strict=True
        ):
            if _CELL.search(self.data, start, end):
                return number
        return 1

    @cached_property
    def max_column(self) -> int:
        letters = {match.group(1) for match in _CELL_COLUMN.finditer(self.data)}
        return max((column_index_from_string(letter.decode()) for letter in letters), default=1)


def split_sheet(data: bytes) -> tuple[bytes, SheetSource] | None:
    """The sheet without its rows, and the rows; None for a layout this does not handle.

    Shared formulas are left out because a row's formula can depend on the row that
    declares it, and rows without a number because they are numbered by their position.
    """
    namespace = _NAMESPACE.search(data)
    start = data.find(_SHEET_DATA)
    end = data.find(_SHEET_DATA_END)
    if namespace is None or start < 0 or end < start or b't="shared"' in data:
        return None
    numbers: list[int] = []
    spans: list[tuple[int, int]] = []
    hidden: list[int] = []
    starts = list(_ROW.finditer(data, start, end))
    for position, tag in enumerate(starts):
        number = _ROW_NUMBER.search(tag.group())
        if number is None:
            return None
        numbers.append(int(number.group(1)))
        following = starts[position + 1].start() if position + 1 < len(starts) else end
        spans.append((tag.start(), following))
        if _HIDDEN.search(tag.group()):
            hidden.append(numbers[-1])
    if numbers != sorted(set(numbers)):
        return None
    without_rows = data[:start] + _SHEET_DATA + b"/>" + data[end + len(_SHEET_DATA_END) :]
    return without_rows, SheetSource(data, namespace.group(1), numbers, spans, hidden)


class LazyCells(dict[CellKey, Any]):
    """The cells of a worksheet, parsed from its rows the first time one is asked for."""

    def __init__(self, sheet: Worksheet, source: SheetSource, strings: list[str]) -> None:
        super().__init__()
        self.sheet = sheet
        self.source = source
        self.strings = strings
        self.loaded: set[int] = set()
        self.columns: dict[int, list[int]] = {}

    def keep(self, cell: Any) -> None:
        self[(cell.row, cell.column)] = cell
        self.columns.setdefault(cell.row, []).append(cell.column)

    def get(self, key: CellKey, default: Any = None) -> Any:
        self._load(key[0])
        return super().get(key, default)

    def __getitem__(self, key: CellKey) -> Any:
        self._load(key[0])
        return super().__getitem__(key)

    def __contains__(self, key: object) -> bool:
        self._load(key[0])  # pyright: ignore[reportIndexIssue]
        return super().__contains__(key)

    def stored_in(
        self, top: int, left: int, bottom: int, right: int
    ) -> Iterator[tuple[CellKey, Any]]:
        for row in self.source.row_numbers[self.source.between(top, bottom)]:
            self._load(row)
            for col in sorted(self.columns.get(row, ())):
                if left <= col <= right:
                    yield (row, col), super().__getitem__((row, col))

    def _load(self, row: int) -> None:
        block = row // BLOCK_ROWS
        if block in self.loaded:
            return
        self.loaded.add(block)
        spans = self.source.spans[
            self.source.between(block * BLOCK_ROWS, (block + 1) * BLOCK_ROWS - 1)
        ]
        if not spans:
            return
        rows = b"".join(self.source.data[start:end] for start, end in spans)
        document = b'<worksheet xmlns="%s"><sheetData>%s</sheetData></worksheet>' % (
            self.source.namespace,
            rows,
        )
        book: Any = self.sheet.parent
        parser = WorkSheetParser(
            BytesIO(document),
            self.strings,
            book._data_only,
            book.epoch,
            book._date_formats,
            book._timedelta_formats,
        )
        for _, parsed in parser.parse():
            for item in parsed:
                key = (item["row"], item["column"])
                if dict.__contains__(self, key):
                    continue
                cell = Cell(
                    self.sheet,
                    row=key[0],
                    column=key[1],
                    style_array=book._cell_styles[item["style_id"]],
                )
                cell._value = item["value"]
                cell.data_type = item["data_type"]
                self.keep(cell)


Parts = dict[str, tuple[bytes, SheetSource | None]]


class _SheetArchive(ZipFile):
    """An archive that serves each worksheet without its rows, which are kept in ``parts``."""

    def __init__(self, path: Path, parts: Parts, sheet_parts: set[str]) -> None:
        super().__init__(path)
        self.parts = parts
        self.sheet_parts = sheet_parts
        self.worksheets: list[str] = []

    def open(self, name: Any, mode: Any = "r", pwd: Any = None, **options: Any) -> IO[bytes]:
        name = getattr(name, "filename", name)
        if name not in self.parts:
            with super().open(name, mode, pwd, **options) as part:
                data = part.read()
            if name not in self.sheet_parts:
                return BytesIO(data)
            split = split_sheet(data)
            self.parts[name] = split if split else (data, None)
        self.worksheets.append(name)
        return BytesIO(self.parts[name][0])


def load_lazy(path: Path, parts: Parts, *, data_only: bool) -> Workbook:
    """Load a workbook the way ``load_workbook`` does, leaving the cells to be read on demand.

    A sheet this cannot split into rows is loaded whole. ``parts`` is shared by the loads of
    one file, so that each sheet is read from the file once.
    """
    reader = ExcelReader(path, data_only=data_only, keep_links=False)
    reader.read_manifest()
    sheet_parts = {part.PartName.lstrip("/") for part in reader.package.findall(WORKSHEET_TYPE)}
    reader.archive.close()
    archive = reader.archive = _SheetArchive(path, parts, sheet_parts)
    reader.read()
    sheets = [sheet for sheet in reader.wb.worksheets if isinstance(sheet, Worksheet)]
    for sheet, name in zip(sheets, archive.worksheets, strict=True):
        source = parts[name][1]
        if source is None:
            continue
        cells = LazyCells(sheet, source, reader.shared_strings)
        for cell in sheet._cells.values():
            if isinstance(cell, MergedCell):
                cells.keep(cell)
        sheet._cells = cells
        for row in source.hidden_rows:
            sheet.row_dimensions[row].hidden = True
    return reader.wb


def extent(sheet: Worksheet) -> tuple[int, int]:
    """The last row and column that hold a cell, which loading a lazy sheet's rows must not."""
    cells = sheet._cells
    if isinstance(cells, LazyCells):
        return cells.source.max_row, cells.source.max_column
    return sheet.max_row, sheet.max_column
