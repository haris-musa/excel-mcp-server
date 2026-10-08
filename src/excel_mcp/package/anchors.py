"""Where shapes sit on a sheet, and how an edit of rows or columns moves them.

A shape is anchored to cells by two corners (``from`` and ``to``, each a cell and an offset
into it in EMU). Excel moves and resizes a shape that is anchored to both corners, moves one
that is anchored to its first corner (``oneCell``), and leaves an absolute one where it is.
"""

import re
from dataclasses import dataclass, replace

from openpyxl.worksheet.worksheet import Worksheet

from excel_mcp.package.lines import LineEdit

EMU_PER_POINT = 12700
EMU_PER_PIXEL = 9525
_DEFAULT_COLUMN_PIXELS = 64

_ANCHOR = re.compile(
    r"<(?P<name>(?:\w+:)?(?:twoCellAnchor|oneCellAnchor|anchor))\b(?P<attrs>[^>]*)>"
    r"(?P<body>.*?)</(?P=name)>",
    re.S,
)
_CORNER = re.compile(
    r"<(?P<p>(?:\w+:)?)(?P<tag>from|to)>\s*"
    r"<(?P<q>(?:\w+:)?)col>(?P<col>\d+)</(?P=q)col>\s*"
    r"<(?P=q)colOff>(?P<col_off>-?\d+)</(?P=q)colOff>\s*"
    r"<(?P=q)row>(?P<row>\d+)</(?P=q)row>\s*"
    r"<(?P=q)rowOff>(?P<row_off>-?\d+)</(?P=q)rowOff>\s*"
    r"</(?P=p)(?P=tag)>"
)
_SHAPE_ID = re.compile(r'<(?:\w+:)?cNvPr\b[^>]*?\bid="(\d+)"')


@dataclass(frozen=True)
class Corner:
    """A cell (0-based) and the offsets into it, in EMU."""

    col: int
    col_off: int
    row: int
    row_off: int


Placement = tuple[Corner, Corner | None]
Moved = dict[int, Placement]


class Geometry:
    """Sizes of the rows and columns of a sheet before an edit, in EMU."""

    def __init__(self, sheet: Worksheet) -> None:
        self.sheet = sheet
        self.default_row = round((sheet.sheet_format.defaultRowHeight or 15) * EMU_PER_POINT)
        width = sheet.sheet_format.defaultColWidth
        self.default_column = (_pixels(width) if width else _DEFAULT_COLUMN_PIXELS) * EMU_PER_PIXEL
        self.columns: dict[int, int] = {}
        for dimension in sheet.column_dimensions.values():
            if dimension.min is None or dimension.max is None:
                continue
            size = 0 if dimension.hidden else self._column(dimension.width)
            self.columns.update(dict.fromkeys(range(dimension.min, dimension.max + 1), size))

    def size(self, rows: bool, line: int) -> int:
        """The size of a line (1-based)."""
        if not rows:
            return self.columns.get(line, self.default_column)
        dimension = self.sheet.row_dimensions.get(line)
        if dimension is None:
            return self.default_row
        if dimension.hidden:
            return 0
        return round(dimension.height * EMU_PER_POINT) if dimension.height else self.default_row

    def _column(self, width: float | None) -> int:
        return self.default_column if width is None else _pixels(width) * EMU_PER_PIXEL


def _pixels(width: float) -> int:
    return int((256 * width + int(128 / 7)) / 256 * 7)


def move_anchors(xml: str, edit: LineEdit, geometry: Geometry, moved: Moved) -> str:
    """The XML with every anchor in it moved as Excel moves the shapes of an edited sheet.

    ``moved`` collects the new place of each shape by its id, for the VML that repeats it.
    """

    def rewrite(match: re.Match[str]) -> str:
        how = _placement(match["name"], match["attrs"])
        if how == "absolute":
            return match[0]
        corners = {m["tag"]: _corner(m) for m in _CORNER.finditer(match["body"])}
        new_first, new_last = _move(corners["from"], corners.get("to"), how, edit, geometry)
        body = _CORNER.sub(
            lambda m: _text(m, new_first if m["tag"] == "from" else new_last), match["body"]
        )
        if (shape := _SHAPE_ID.search(body)) is not None:
            moved[int(shape[1])] = (new_first, new_last)
        return match[0].replace(match["body"], body, 1)

    return _ANCHOR.sub(rewrite, xml)


def _placement(name: str, attributes: str) -> str:
    """How the shape follows its cells: ``twoCell``, ``oneCell`` or ``absolute``."""
    if name.endswith("oneCellAnchor"):
        return "oneCell"
    if name.endswith("twoCellAnchor"):
        found = re.search(r'\beditAs="(\w+)"', attributes)
        return found[1] if found else "twoCell"
    if 'sizeWithCells="1"' in attributes:
        return "twoCell"
    return "oneCell" if 'moveWithCells="1"' in attributes else "absolute"


def _move(
    first: Corner, last: Corner | None, how: str, edit: LineEdit, geometry: Geometry
) -> Placement:
    rows = edit.axis == "rows"
    line, offset = (first.row, first.row_off) if rows else (first.col, first.col_off)
    new_line, reset = edit.corner(line + 1)
    new_first = _with(first, rows, new_line - 1, 0 if reset else offset)
    if last is None:
        return new_first, None
    end, end_offset = (last.row, last.row_off) if rows else (last.col, last.col_off)
    if how == "twoCell":
        new_end, lost = edit.corner(end + 1)
        return new_first, _with(last, rows, new_end - 1, 0 if lost else end_offset)
    size = sum(geometry.size(rows, n + 1) for n in range(line, end)) + end_offset - offset
    return new_first, _after(
        last, rows, edit, geometry, new_line - 1, size + (0 if reset else offset)
    )


def _after(
    last: Corner, rows: bool, edit: LineEdit, geometry: Geometry, line: int, distance: int
) -> Corner:
    """The corner ``distance`` EMU after the top of ``line``, on the sheet as it is edited."""
    while True:
        size = geometry.size(rows, edit.source(line + 1))
        if distance < size:
            return _with(last, rows, line, distance)
        distance -= size
        line += 1


def _with(corner: Corner, rows: bool, line: int, offset: int) -> Corner:
    if rows:
        return replace(corner, row=line, row_off=offset)
    return replace(corner, col=line, col_off=offset)


def _corner(match: re.Match[str]) -> Corner:
    return Corner(
        int(match["col"]), int(match["col_off"]), int(match["row"]), int(match["row_off"])
    )


def _text(match: re.Match[str], corner: Corner | None) -> str:
    if corner is None:
        return match[0]
    q, p, tag = match["q"], match["p"], match["tag"]
    return (
        f"<{p}{tag}><{q}col>{corner.col}</{q}col><{q}colOff>{corner.col_off}</{q}colOff>"
        f"<{q}row>{corner.row}</{q}row><{q}rowOff>{corner.row_off}</{q}rowOff></{p}{tag}>"
    )


_VML_ANCHOR = re.compile(r"(<x:Anchor>)([^<]*)(</x:Anchor>)")
_VML_ID = re.compile(r'(?:\bspid|\bid)="_x0000_s(\d+)"')
_VML_CELL = re.compile(r"<x:(Row|Column)>(\d+)</x:\1>")


def move_vml_shape(shape: str, edit: LineEdit, moved: Moved) -> str:
    """A VML shape whose anchor follows the drawing's shape with the same id, or moves like a
    two-cell anchor when the drawing has none. Anchors there are in pixels."""
    found = _VML_ID.search(shape)
    placed = moved.get(int(found[1])) if found else None

    def anchor(match: re.Match[str]) -> str:
        if placed is not None:
            return f"{match[1]}{_vml_anchor(placed)}{match[3]}"
        numbers = [int(n) for n in match[2].split(",")]
        return f"{match[1]}{_moved_numbers(numbers, edit)}{match[3]}"

    def cell(match: re.Match[str]) -> str:
        if (match[1] == "Row") != (edit.axis == "rows"):
            return match[0]
        line, _ = edit.corner(int(match[2]) + 1)
        return f"<x:{match[1]}>{line - 1}</x:{match[1]}>"

    return _VML_CELL.sub(cell, _VML_ANCHOR.sub(anchor, shape))


def _vml_anchor(placed: Placement) -> str:
    first, last = placed
    corners = [c for c in (first, last) if c is not None]
    numbers = [
        n
        for c in corners
        for n in (c.col, c.col_off // EMU_PER_PIXEL, c.row, c.row_off // EMU_PER_PIXEL)
    ]
    return ", ".join(str(n) for n in numbers)


def _moved_numbers(numbers: list[int], edit: LineEdit) -> str:
    """Left column, offset, top row, offset, right column, offset, bottom row, offset."""
    rows = edit.axis == "rows"
    for index in (2, 6) if rows else (0, 4):
        line, lost = edit.corner(numbers[index] + 1)
        numbers[index] = line - 1
        if lost:
            numbers[index + 1] = 0
    return ", ".join(str(n) for n in numbers)
