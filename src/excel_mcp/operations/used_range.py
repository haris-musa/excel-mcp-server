"""The used range of a sheet, found in its XML text without parsing the cells.

Parsing every cell of a big sheet takes seconds; the text of the file holds the answer in a
form a regular expression finds at the speed of a copy. Only sheets laid out the way Excel and
openpyxl write them are read this way: every row and cell starts with its own reference
(``<row r="5"``, ``<c r="B5"``) and rows are in order. For any other layout the caller
parses the cells, which gives the same result.
"""

import re

from openpyxl.utils import column_index_from_string

from excel_mcp.refs import CellRange

_ROW = re.compile(rb'<row r="(\d+)"')
# A cell that holds a value or a formula: its first child has content ("<v></v>" has none).
_VALUE = re.compile(rb'<c r="([A-Z]+)\d+"[^>/]*>\s*(?:<f[\s>/]|<v>[^<]|<is[\s>/])')
_EMPTY_DATA = re.compile(rb"<sheetData\s*/>")
_EMPTY = CellRange(1, 1, 1, 1)


def used_range_in(xml: bytes) -> CellRange | None:
    """The smallest range holding every cell with a value, or None for an unusual layout."""
    if _EMPTY_DATA.search(xml):
        return _EMPTY
    start, end = xml.find(b"<sheetData"), xml.find(b"</sheetData>")
    if start < 0 or end < start:
        return None
    data = xml[start:end]
    rows = [(int(match[1]), match.start()) for match in _ROW.finditer(data)]
    numbers = [number for number, _ in rows]
    if (
        data.count(b"<row") != len(rows)
        or numbers != sorted(set(numbers))
        or data.count(b"<c ") + data.count(b"<c>") != data.count(b'<c r="')
    ):
        return None
    starts = [offset for _, offset in rows] + [len(data)]
    spans = [(rows[i][0], starts[i], starts[i + 1]) for i in range(len(rows))]
    first = next((number for number, a, b in spans if _VALUE.search(data, a, b)), None)
    if first is None:
        return _EMPTY
    last = next(number for number, a, b in reversed(spans) if _VALUE.search(data, a, b))
    columns = {column_index_from_string(letters.decode()) for letters in set(_VALUE.findall(data))}
    return CellRange(first, min(columns), last, max(columns))
