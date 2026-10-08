"""The pivot cache: a snapshot of the source data that a PivotTable reads from."""

import datetime as dt
from dataclasses import dataclass

from openpyxl.pivot.cache import (
    CacheDefinition,
    CacheField,
    CacheSource,
    SharedItems,
    WorksheetSource,
)
from openpyxl.pivot.fields import DateTimeField, Index, Missing, Number, Text
from openpyxl.pivot.record import Record, RecordList
from openpyxl.utils.datetime import to_excel

from excel_mcp.operations.pivot_source import Column, Source, Value

CacheItem = Missing | Number | Text | DateTimeField


@dataclass(frozen=True)
class FieldItems:
    """The distinct values of a field used to group.

    ``shared`` holds each value once, in order of first appearance, which is how the
    cache stores them. ``index`` gives each record's value as a position in ``shared``.
    A PivotTable lists the values sorted, blanks last: ``order`` is that listing as
    positions in ``shared`` and ``rank`` the inverse.
    """

    shared: list[Value]
    index: list[int]
    order: list[int]
    rank: list[int]

    def label(self, rank: int) -> Value:
        return self.shared[self.order[rank]]

    def record_rank(self, record: int) -> int:
        return self.rank[self.index[record]]


def build_items(column: Column) -> FieldItems:
    seen: dict[object, int] = {}
    shared: list[Value] = []
    index: list[int] = []
    for value in column.values:
        key = value.casefold() if isinstance(value, str) else value
        if key not in seen:
            seen[key] = len(shared)
            shared.append(value)
        index.append(seen[key])
    order = sorted(range(len(shared)), key=lambda item: _sort_key(shared[item]))
    rank = [0] * len(shared)
    for position, item in enumerate(order):
        rank[item] = position
    return FieldItems(shared, index, order, rank)


def _sort_key(value: Value) -> tuple[bool, object]:
    if value is None:
        return True, 0
    return False, value.casefold() if isinstance(value, str) else value


def build_cache(source: Source, items: dict[int, FieldItems]) -> CacheDefinition:
    """Describe ``source`` with shared items for the fields in ``items``."""
    cache = CacheDefinition(
        cacheSource=CacheSource(
            type="worksheet",
            worksheetSource=WorksheetSource(ref=source.ref, sheet=source.sheet),
        ),
        cacheFields=[
            CacheField(
                name=column.name,
                numFmtId=column.number_format_id,
                sharedItems=_shared_items(column, items.get(position)),
            )
            for position, column in enumerate(source.columns)
        ],
        refreshedBy="excel-mcp-server",
        refreshedDate=to_excel(dt.datetime.now(dt.UTC).replace(tzinfo=None)),
        createdVersion=6,
        refreshedVersion=6,
        minRefreshableVersion=3,
        recordCount=source.record_count,
    )
    cache.records = RecordList(r=_records(source, items))
    return cache


def _shared_items(column: Column, items: FieldItems | None) -> SharedItems:
    present = [value for value in column.values if value is not None]
    blank = len(present) < len(column.values)
    numbers = [float(value) for value in present if isinstance(value, int | float)]
    moments = [value for value in present if isinstance(value, dt.datetime)]
    is_number, is_date = column.kind == "number", column.kind == "date"
    return SharedItems(
        _fields=[_item(value) for value in items.shared] if items else (),
        containsBlank=True if blank else None,
        containsSemiMixedTypes=False if column.kind != "text" and not blank else None,
        containsString=False if column.kind != "text" else None,
        containsNumber=True if is_number else None,
        containsInteger=all(number.is_integer() for number in numbers) if is_number else None,
        minValue=min(numbers, default=None),
        maxValue=max(numbers, default=None),
        containsDate=True if is_date else None,
        containsNonDate=False if is_date else None,
        minDate=min(moments, default=None),
        maxDate=max(moments, default=None),
    )


def _item(value: Value) -> CacheItem:
    match value:
        case None:
            return Missing()
        case str():
            return Text(v=value)
        case dt.datetime():
            return DateTimeField(v=value)
        case _:
            return Number(v=value)


def _records(source: Source, items: dict[int, FieldItems]) -> list[Record]:
    return [
        Record(
            _fields=[
                Index(v=items[position].index[row])
                if position in items
                else _item(column.values[row])
                for position, column in enumerate(source.columns)
            ]
        )
        for row in range(source.record_count)
    ]
