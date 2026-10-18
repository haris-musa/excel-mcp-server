"""Workbooks whose cells are parsed, row by row, when something asks for them.

Calculating a page of a large sheet needs the cells of that page and of what its formulas
use, not the other million. The workbook is loaded without its cell data, and each sheet
keeps its XML so rows can be parsed on demand with openpyxl's own cell parser. Only the
cells and the hidden state of rows are loaded this way: that is all the calculator reads.
`lazy_edit` builds on this to edit large sheets without parsing the rows an edit leaves alone.
"""

import re
from bisect import bisect_left, bisect_right
from collections.abc import Callable, Iterator
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
from openpyxl.xml.constants import SHEET_MAIN_NS, WORKSHEET_TYPE

BLOCK_ROWS = 128
_SHEET_DATA = b"<sheetData"
_SHEET_DATA_END = b"</sheetData>"
_ROW = re.compile(rb"<row\b[^>]*>")
_ROW_NUMBER = re.compile(rb'\br="(\d+)"')
_HIDDEN = re.compile(rb'\bhidden="(?:1|true)"')
_CELL = re.compile(rb"<c[ >/]")
_CELL_COLUMN = re.compile(rb'<c\b[^>]*?\br="([A-Z]+)\d')
_ROOT = re.compile(rb'<worksheet\b[^>]*?\bxmlns="([^"]+)"[^>]*>')

CellKey = tuple[int, int]


@dataclass
class SheetSource:
    """A sheet's XML and where each of its rows is in it."""

    data: bytes
    root: bytes
    row_numbers: list[int]
    spans: list[tuple[int, int]]
    hidden_rows: list[int]

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

    Shared formulas are left out because a row's formula can depend on the row that
    declares it, and rows without a number because they are numbered by their position.
    """
    root = _ROOT.search(data)
    start = data.find(_SHEET_DATA)
    end = data.find(_SHEET_DATA_END)
    if (
        root is None
        or root[1] != SHEET_MAIN_NS.encode()
        or start < 0
        or end < start
        or b't="shared"' in data
    ):
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
    return without_rows, SheetSource(data, root.group(), numbers, spans, hidden)


class LazyCells(dict[CellKey, Any]):
    """The cells of a worksheet, parsed from its rows the first time one is asked for.

    Asking for one cell parses the block of rows around it; asking for all of them (iterating,
    counting) parses the sheet. Changing a cell parses its block first, so a block that was
    parsed holds all of its rows: what is not in ``loaded`` is as the file has it.
    """

    def __init__(self, sheet: Worksheet, source: SheetSource, strings: list[str]) -> None:
        super().__init__()
        self.sheet = sheet
        self.source = source
        self.strings = strings
        self.loaded: set[int] = set()
        self.columns: dict[int, list[int]] = {}
        self.last_row = self.last_column = 0

    def keep(self, cell: Any) -> None:
        dict.__setitem__(self, (cell.row, cell.column), cell)
        self.columns.setdefault(cell.row, []).append(cell.column)

    def held(self) -> dict[CellKey, Any]:
        """The cells parsed or put in so far."""
        return dict(dict.items(self))

    def is_loaded(self, row: int) -> bool:
        return row // BLOCK_ROWS in self.loaded

    def extent(self) -> tuple[int, int]:
        """The last row and column that hold a cell, without parsing any."""
        return max(self.source.max_row, self.last_row), max(
            self.source.max_column, self.last_column
        )

    def get(self, key: CellKey, default: Any = None) -> Any:
        self.load(key[0])
        return super().get(key, default)

    def __getitem__(self, key: CellKey) -> Any:
        self.load(key[0])
        return super().__getitem__(key)

    def __contains__(self, key: object) -> bool:
        self.load(key[0])  # pyright: ignore[reportIndexIssue]
        return super().__contains__(key)

    def __setitem__(self, key: CellKey, cell: Any) -> None:
        self.load(key[0])
        if not super().__contains__(key):
            self.columns.setdefault(key[0], []).append(key[1])
        super().__setitem__(key, cell)
        self.last_row, self.last_column = max(self.last_row, key[0]), max(self.last_column, key[1])

    def __delitem__(self, key: CellKey) -> None:
        self.load(key[0])
        super().__delitem__(key)
        self.columns[key[0]].remove(key[1])

    def pop(self, key: CellKey, *default: Any) -> Any:
        self.load(key[0])
        if super().__contains__(key):
            self.columns[key[0]].remove(key[1])
        return super().pop(key, *default)

    def setdefault(self, key: CellKey, default: Any = None) -> Any:
        if key not in self:
            self[key] = default
        return super().__getitem__(key)

    def __bool__(self) -> bool:
        return dict.__len__(self) > 0 or self.source.has_cells

    # Everything that looks at all the cells parses them all.

    def __iter__(self) -> Iterator[CellKey]:
        self.load_all()
        return super().__iter__()

    def __reversed__(self) -> Iterator[CellKey]:
        self.load_all()
        return super().__reversed__()

    def __len__(self) -> int:
        self.load_all()
        return super().__len__()

    def keys(self) -> Any:
        self.load_all()
        return super().keys()

    def values(self) -> Any:
        self.load_all()
        return super().values()

    def items(self) -> Any:
        self.load_all()
        return super().items()

    def copy(self) -> dict[CellKey, Any]:
        self.load_all()
        return super().copy()

    def stored_in(
        self, top: int, left: int, bottom: int, right: int
    ) -> Iterator[tuple[CellKey, Any]]:
        for row in self.source.row_numbers[self.source.between(top, bottom)]:
            self.load(row)
            for col in sorted(self.columns.get(row, ())):
                if left <= col <= right:
                    yield (row, col), super().__getitem__((row, col))

    def load_all(self) -> None:
        source = self.source
        spans = [
            span
            for number, span in zip(source.row_numbers, source.spans, strict=True)
            if number // BLOCK_ROWS not in self.loaded
        ]
        self.loaded.update(number // BLOCK_ROWS for number in source.row_numbers)
        self._parse(spans)

    def load(self, row: int) -> None:
        """Parse the block of rows that holds ``row``, unless that is done."""
        block = row // BLOCK_ROWS
        if block in self.loaded:
            return
        self.loaded.add(block)
        first, last = block * BLOCK_ROWS, (block + 1) * BLOCK_ROWS - 1
        self._parse(self.source.spans[self.source.between(first, last)])

    def _parse(self, spans: list[tuple[int, int]]) -> None:
        if not spans:
            return
        rows = b"".join(self.source.data[start:end] for start, end in spans)
        document = self.source.root + b"<sheetData>" + rows + b"</sheetData></worksheet>"
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
                if super().__contains__(key):
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
Splitter = Callable[[bytes], tuple[bytes, SheetSource] | None]


class _SheetArchive(ZipFile):
    """An archive that serves each worksheet without its rows, which are kept in ``parts``."""

    def __init__(self, path: Path, parts: Parts, sheet_parts: set[str], split: Splitter) -> None:
        super().__init__(path)
        self.parts = parts
        self.sheet_parts = sheet_parts
        self.split = split
        self.worksheets: list[str] = []

    def open(self, name: Any, mode: Any = "r", pwd: Any = None, **options: Any) -> IO[bytes]:
        name = getattr(name, "filename", name)
        if name not in self.parts:
            with super().open(name, mode, pwd, **options) as part:
                data = part.read()
            if name not in self.sheet_parts:
                return BytesIO(data)
            split = self.split(data)
            self.parts[name] = split if split else (data, None)
        self.worksheets.append(name)
        return BytesIO(self.parts[name][0])


def open_lazily(path: Path, parts: Parts, split: Splitter, **options: Any) -> ExcelReader:
    """A reader that has not read the workbook yet, whose sheets are served by ``split``."""
    reader = ExcelReader(path, **options)
    reader.read_manifest()
    sheet_parts = {part.PartName.lstrip("/") for part in reader.package.findall(WORKSHEET_TYPE)}
    reader.archive.close()
    reader.archive = _SheetArchive(path, parts, sheet_parts, split)
    return reader


def lazy_sheets(reader: ExcelReader, parts: Parts) -> Iterator[tuple[Worksheet, SheetSource]]:
    """The worksheets of a workbook read by ``reader`` whose rows were split off."""
    archive: Any = reader.archive
    sheets = [sheet for sheet in reader.wb.worksheets if isinstance(sheet, Worksheet)]
    for sheet, name in zip(sheets, archive.worksheets, strict=True):
        if (source := parts[name][1]) is not None:
            yield sheet, source


def load_lazy(path: Path, parts: Parts, *, data_only: bool) -> Workbook:
    """Load a workbook the way ``load_workbook`` does, leaving the cells to be read on demand.

    A sheet this cannot split into rows is loaded whole. ``parts`` is shared by the loads of
    one file, so that each sheet is read from the file once.
    """
    reader = open_lazily(path, parts, split_sheet, data_only=data_only, keep_links=False)
    reader.read()
    for sheet, source in lazy_sheets(reader, parts):
        cells = LazyCells(sheet, source, reader.shared_strings)
        for cell in sheet._cells.values():
            if isinstance(cell, MergedCell):
                cells.keep(cell)
        sheet._cells = cells
        for row in source.hidden_rows:
            sheet.row_dimensions[row].hidden = True
    return reader.wb


def held_cells(sheet: Worksheet) -> dict[CellKey, Any]:
    """The cells a sheet holds in memory: all of them, but for the rows a lazy sheet has not parsed.

    For looking for what can only be in cells that were parsed because it matters to them
    (hyperlinks, comments, array formulas) or put in by an edit.
    """
    cells = sheet._cells
    return cells.held() if isinstance(cells, LazyCells) else cells


def extent(sheet: Worksheet) -> tuple[int, int]:
    """The last row and column that hold a cell, which loading a lazy sheet's rows must not."""
    cells = sheet._cells
    if isinstance(cells, LazyCells):
        return cells.extent()
    return sheet.max_row, sheet.max_column
