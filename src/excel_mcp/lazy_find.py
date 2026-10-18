"""Finding which rows of a sheet that is not parsed can hold a text, without parsing them.

Replacing text looks at every cell, but most of a large sheet holds none of it. Only the blocks
of rows that might hold a match are parsed; the others are as in the file. "Might" is decided
from the XML and errs on the side of parsing: a block that holds nothing is parsed at worst.
A cell is searched as `replace` searches it: text, numbers as their text, formulas as they are
written (without the storage prefixes), so a formula that is only a master of the shared
formulas below it matches for the cells that follow it too.
"""

import re
from bisect import bisect_right
from collections.abc import Callable
from html import unescape

from excel_mcp.lazy_workbook import LazyCells

_STRING = re.compile(rb'<c\b[^>]*\bt="s"[^>]*><v>(\d+)</v>')
_INLINE = re.compile(rb"<is>(.*?)</is>", re.S)
_TAGS = re.compile(rb"<[^>]*>")
_FORMULA = re.compile(rb"<f\b([^>]*)>([^<]*)</f>")
_NUMBER = re.compile(rb'<c\b(?![^>]*\bt="(?!n")\w+")[^>]*><v>')
_GROUP = re.compile(rb'\bsi="(\d+)"')
_TEXT_OF_NUMBERS = frozenset("0123456789.-+e")
# What a stored formula has and its text as shown does not, or the reverse.
_STORAGE_MARKS = ("_xl", "anchorarray")


def load_matches(
    cells: LazyCells, matches: Callable[[str], bool], query: str, *, in_formulas: bool
) -> None:
    """Parse the blocks that may hold a cell for which ``matches(text)`` is true.

    ``query`` is what ``matches`` looks for, which tells what kinds of cell it can match.
    """
    source = cells.source
    starts = [start for start, _ in source.spans]
    rows: set[int] = set()

    def mark(position: int) -> None:
        rows.add(source.row_numbers[bisect_right(starts, position) - 1])

    hits = {index for index, text in enumerate(cells.strings) if matches(text)}
    if hits:
        for found in _STRING.finditer(source.data):
            if int(found[1]) in hits:
                mark(found.start())
    for found in _INLINE.finditer(source.data):
        if matches(unescape(_TAGS.sub(b"", found[1]).decode())):
            mark(found.start())
    if set(query.lower()) <= _TEXT_OF_NUMBERS:
        rows.update(_rows_with_numbers(cells))
    if in_formulas:
        _mark_formulas(cells, matches, query, mark)
    for number in rows:
        cells.load(number)


def _rows_with_numbers(cells: LazyCells) -> list[int]:
    source = cells.source
    return [
        number
        for number, (start, stop) in zip(source.row_numbers, source.spans, strict=True)
        if _NUMBER.search(source.data, start, stop)
    ]


def _mark_formulas(
    cells: LazyCells, matches: Callable[[str], bool], query: str, mark: Callable[[int], None]
) -> None:
    source = cells.source
    digits = any(character.isdigit() for character in query)
    for found in _FORMULA.finditer(source.data):
        shown = "=" + unescape(found[2].decode())
        hit = matches(shown) or any(part in shown.lower() for part in _STORAGE_MARKS)
        if hit:
            mark(found.start())
        group = _GROUP.search(found[1])
        master = source.masters.get(group[1].decode()) if group else None
        # A cell that follows its master differs from it in digits, and in letters when the
        # group spans columns.
        if master is not None and (hit or digits or master.wide):
            cells.load_between(*master.rows)
