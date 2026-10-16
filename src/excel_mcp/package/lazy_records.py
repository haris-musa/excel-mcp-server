"""PivotTable cache records that stay as the bytes of the file until something needs them.

A big source has millions of record items, and openpyxl turns each into an object when it
loads a workbook and again when it saves it: minutes and gigabytes for an edit that never
looks at them. `install` makes openpyxl keep the part as it is and write it back unchanged.

What needs the records, such as giving a field its shared items for a slicer, reads one
column of them (`column`) or replaces one (`share`), both a pass over the XML that holds one
record at a time.
"""

# pyright: reportAttributeAccessIssue=false, reportOptionalMemberAccess=false
# openpyxl's stubs lack the `deps` it sets on what it reads, and type `get_rel` as Optional.

from collections.abc import Iterator
from io import BytesIO
from typing import Any
from xml.sax.saxutils import quoteattr

from openpyxl.packaging.relationship import get_rel
from openpyxl.pivot import fields
from openpyxl.pivot.cache import CacheDefinition
from openpyxl.pivot.record import RecordList
from openpyxl.reader.workbook import WorkbookParser
from openpyxl.xml.constants import SHEET_MAIN_NS
from openpyxl.xml.functions import iterparse

ITEM_TYPES = {
    "m": fields.Missing,
    "n": fields.Number,
    "b": fields.Boolean,
    "e": fields.Error,
    "s": fields.Text,
    "d": fields.DateTimeField,
    "x": fields.Index,
}


class LazyRecords(RecordList):
    """Records held as the XML of the part."""

    def __init__(self, xml: bytes) -> None:
        super().__init__()
        self.xml = xml

    @property
    def count(self) -> int:
        return sum(1 for _ in _records(self.xml))

    def column(self, position: int) -> list[Any]:
        """The items of one field, as openpyxl's item objects, one per record."""
        return [
            ITEM_TYPES[_local(record[position].tag)].from_tree(record[position])
            for record in _records(self.xml)
        ]

    def share(self, position: int, indexes: list[int]) -> None:
        """Replace the items of one field with positions among its shared items."""
        stream = BytesIO()
        stream.write(f'<pivotCacheRecords xmlns="{SHEET_MAIN_NS}" count="{len(indexes)}">'.encode())
        for record, index in zip(_records(self.xml), indexes, strict=True):
            items = [_xml(item) for item in record]
            items[position] = f'<x v="{index}"/>'
            stream.write(f"<r>{''.join(items)}</r>".encode())
        stream.write(b"</pivotCacheRecords>")
        self.xml = stream.getvalue()

    def _write(self, archive, manifest) -> None:
        archive.writestr(self.path[1:], self.xml)
        manifest.append(self)


def install() -> None:
    """Make openpyxl load records as `LazyRecords`. Safe to call again."""
    WorkbookParser.pivot_caches = property(_pivot_caches)


def _pivot_caches(parser: WorkbookParser) -> dict[int, CacheDefinition]:
    caches = {}
    for entry in parser.caches:
        cache = get_rel(parser.archive, parser.rels, id=entry.id, cls=CacheDefinition)
        if cache.deps:
            cache.records = LazyRecords(parser.archive.read(cache.deps.get(cache.id).target))
        caches[entry.cacheId] = cache
    return caches


def _records(xml: bytes) -> Iterator[Any]:
    """Each ``<r>`` element, which is cleared once the caller has gone on to the next."""
    root = None
    for event, element in iterparse(BytesIO(xml), events=("start", "end")):
        if root is None:
            root = element
        elif event == "end" and _local(element.tag) == "r":
            yield element
            root.clear()


def _xml(element: Any) -> str:
    name = _local(element.tag)
    attributes = "".join(f" {_local(key)}={quoteattr(value)}" for key, value in element.items())
    children = "".join(_xml(child) for child in element)
    return f"<{name}{attributes}>{children}</{name}>" if children else f"<{name}{attributes}/>"


def _local(tag: str) -> str:
    return tag.rpartition("}")[2]
