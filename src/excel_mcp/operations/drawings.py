"""The names and the cells of the charts and pictures on a sheet.

Excel names every shape of a drawing ("Chart 1", "Picture 3"). Unlike a position in a list, a
name survives the deletion of other shapes, so tools select charts and pictures by it.
"""

import re
from typing import Any

from openpyxl.chart._chart import ChartBase
from openpyxl.drawing.image import Image
from openpyxl.drawing.spreadsheet_drawing import OneCellAnchor, TwoCellAnchor
from openpyxl.worksheet.worksheet import Worksheet

from excel_mcp.errors import InvalidArgumentError
from excel_mcp.package.anchors import Geometry
from excel_mcp.package.shape_names import name_of, set_name
from excel_mcp.refs import MAX_COLUMN, MAX_ROW, CellRange
from excel_mcp.text import quoted

EMU_PER_CM = 360_000
EMU_PER_PIXEL = 9525
MAX_NAME_LENGTH = 255
_CORNER = (
    r"<(?:\w+:)?{tag}>\s*<(?:\w+:)?col>(\d+)</(?:\w+:)?col>\s*<(?:\w+:)?colOff>(-?\d+)"
    r"</(?:\w+:)?colOff>\s*<(?:\w+:)?row>(\d+)</(?:\w+:)?row>\s*<(?:\w+:)?rowOff>(-?\d+)<"
)
_FROM = re.compile(_CORNER.format(tag="from"))
_TO = re.compile(_CORNER.format(tag="to"))


def name_uniquely(shapes: list[ChartBase] | list[Image], prefix: str, taken: set[str]) -> None:
    """Give each shape its own name: the one it has, unless taken, else the next free one."""
    for shape in shapes:
        name = name_of(shape)
        if not name or name.casefold() in taken:
            name = free_name(prefix, taken)
        set_name(shape, name)
        taken.add(name.casefold())


def free_name(prefix: str, taken: set[str]) -> str:
    number = 1
    while f"{prefix} {number}".casefold() in taken:
        number += 1
    return f"{prefix} {number}"


def check_name(name: str, taken: set[str]) -> str:
    if not name.strip() or len(name) > MAX_NAME_LENGTH:
        raise InvalidArgumentError(f"Names must be 1 to {MAX_NAME_LENGTH} characters long.")
    if name.casefold() in taken:
        raise InvalidArgumentError(f"A chart or image named {name!r} already exists.")
    return name


def find(names: list[str], name: str, kind: str, sheet: Worksheet) -> int:
    """The 0-based position of `name` in `names`."""
    for position, candidate in enumerate(names):
        if candidate.casefold() == name.strip().casefold():
            return position
    if not names:
        raise InvalidArgumentError(f"Sheet {sheet.title!r} has no {kind}s.")
    raise InvalidArgumentError(
        f"Sheet {sheet.title!r} has no {kind} named {name!r}. Available: {quoted(names)}."
    )


def covered(sheet: Worksheet, anchor: Any) -> str | None:
    """The cells a shape covers, None for one that is not tied to cells."""
    if isinstance(anchor, TwoCellAnchor):
        first, last = anchor._from, anchor.to
        return _between(first.col, first.row, last.col, last.row, last.colOff, last.rowOff)
    if isinstance(anchor, OneCellAnchor):
        first, size = anchor._from, anchor.ext
        return _extent(
            sheet,
            first.col + 1,
            first.row + 1,
            first.colOff + size.width,
            first.rowOff + size.height,
        )
    return None


def extent(sheet: Worksheet, row: int, col: int, width: int, height: int) -> str:
    """The cells a shape of `width` x `height` EMU covers when its corner is at (row, col)."""
    return _extent(sheet, col, row, width, height)


def anchored_range(xml: str) -> str | None:
    """The cells between the corners of a two-cell anchor's XML."""
    first, last = _FROM.search(xml), _TO.search(xml)
    if first is None or last is None:
        return None
    col, _, row, _ = (int(n) for n in first.groups())
    end_col, end_col_off, end_row, end_row_off = (int(n) for n in last.groups())
    return _between(col, row, end_col, end_row, end_col_off, end_row_off)


def _between(col: int, row: int, end_col: int, end_row: int, col_off: int, row_off: int) -> str:
    """The cells from the 0-based corner (`row`, `col`) to the one the offsets lie in."""
    end_col += 1 if col_off or end_col == col else 0
    end_row += 1 if row_off or end_row == row else 0
    return str(CellRange(row + 1, col + 1, end_row, end_col))


def _extent(sheet: Worksheet, col: int, row: int, right: int, bottom: int) -> str:
    """The cells from (`row`, `col`) to the point `right` and `bottom` EMU from their corner."""
    geometry = Geometry(sheet)
    last_row = _last(geometry, True, row, bottom, MAX_ROW)
    last_col = _last(geometry, False, col, right, MAX_COLUMN)
    return str(CellRange(row, col, last_row, last_col))


def _last(geometry: Geometry, rows: bool, line: int, distance: int, limit: int) -> int:
    while line < limit and distance > geometry.size(rows, line):
        distance -= geometry.size(rows, line)
        line += 1
    return line
