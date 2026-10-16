"""The pivot cache records, written straight to XML.

A big source has millions of records items; one openpyxl object per item takes minutes and
gigabytes, so the records are written as text instead.
"""

import datetime as dt
from collections.abc import Iterator, Sequence
from typing import IO

from openpyxl.compat import safe_string
from openpyxl.pivot.record import RecordList
from openpyxl.xml.constants import SHEET_MAIN_NS

from excel_mcp.calc.values import ExcelError
from excel_mcp.operations.pivot_source import Value

ESCAPES = str.maketrans(
    {
        "&": "&amp;",
        "<": "&lt;",
        ">": "&gt;",
        '"': "&quot;",
        "\n": "&#10;",
        "\r": "&#13;",
        "\t": "&#9;",
    }
)
BATCH = 10_000


class SharedIndex(int):
    """The position of a value among its field's shared items."""


Item = Value | SharedIndex
Row = tuple[Item, ...]


class PivotRecords(RecordList):
    """Records that hold their rows as plain values, stored inline; a SharedIndex points at
    a shared item."""

    def __init__(self, rows: Sequence[Row]) -> None:
        super().__init__()
        self.rows = rows

    @property
    def count(self) -> int:
        return len(self.rows)

    def _write(self, archive, manifest) -> None:
        with archive.open(self.path[1:], "w", force_zip64=True) as stream:
            _write_xml(stream, self.rows)
        manifest.append(self)


def _write_xml(stream: IO[bytes], rows: Sequence[Row]) -> None:
    stream.write(f'<pivotCacheRecords xmlns="{SHEET_MAIN_NS}" count="{len(rows)}">'.encode())
    for start in range(0, len(rows), BATCH):
        batch = rows[start : start + BATCH]
        stream.write("".join(f"<r>{''.join(_items(row))}</r>" for row in batch).encode())
    stream.write(b"</pivotCacheRecords>")


def _items(row: Row) -> Iterator[str]:
    for item in row:
        match item:
            case None:
                yield "<m />"
            case SharedIndex():
                yield f'<x v="{int(item)}" />'
            case str():
                yield f'<s v="{item.translate(ESCAPES)}" />'
            case bool():
                yield f'<b v="{int(item)}" />'
            case ExcelError():
                yield f'<e v="{item.code}" />'
            case dt.datetime():
                yield f'<d v="{item.isoformat()}" />'
            case _:
                yield f'<n v="{safe_string(item)}" />'
