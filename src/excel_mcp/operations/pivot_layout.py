"""The layout of a PivotTable: its rows and columns, and the cells Excel shows for them.

The table uses Excel's tabular layout with subtotals for every grouping field but the
innermost, and grand totals. Values sit in columns, below the column fields.
"""

import datetime as dt
from dataclasses import dataclass, replace
from typing import Literal

from openpyxl.pivot.fields import Index
from openpyxl.pivot.table import RowColItem

from excel_mcp.operations.pivot_cache import FieldItems
from excel_mcp.operations.pivot_source import Source, Value

Function = Literal["sum", "count", "average", "min", "max"]
CellContent = str | int | float | dt.datetime | None
CellOut = tuple[CellContent, str | None]
GroupKey = tuple[tuple[int, ...], tuple[int, ...]]

_ITEM_TYPE: dict[str, Literal["data", "default", "grand"]] = {
    "leaf": "data",
    "subtotal": "default",
    "grand": "grand",
}


@dataclass(frozen=True)
class DataSpec:
    field: int
    function: Function
    caption: str


@dataclass(frozen=True)
class Slot:
    """One row of the table body or one column of it."""

    kind: Literal["leaf", "subtotal", "grand"]
    path: tuple[int, ...]
    data: int = 0


@dataclass(frozen=True)
class Shape:
    source: Source
    rows: list[int]
    columns: list[int]
    values: list[DataSpec]
    items: dict[int, FieldItems]

    @property
    def header_rows(self) -> int:
        return 1 + len(self.columns) + (len(self.values) > 1)


def axis_slots(fields: list[int], shape: Shape, data_count: int = 1) -> list[Slot]:
    """The lines of an axis: each distinct group, its subtotals, then the grand total."""
    if not fields:
        return [Slot("leaf", (), data) for data in range(data_count)]
    keys = {_key(shape, fields, record) for record in range(shape.source.record_count)}
    leaves = sorted(keys)
    slots = []
    for position, leaf in enumerate(leaves):
        slots.append(Slot("leaf", leaf))
        following = leaves[position + 1] if position + 1 < len(leaves) else None
        for size in range(len(fields) - 1, 0, -1):
            if following is None or following[:size] != leaf[:size]:
                slots.append(Slot("subtotal", leaf[:size]))
    slots.append(Slot("grand", ()))
    if data_count == 1:
        return slots
    return [replace(slot, data=data) for slot in slots for data in range(data_count)]


def axis_items(slots: list[Slot], with_data: bool) -> list[RowColItem]:
    """Describe the lines of an axis the way a PivotTable definition lists them."""
    items = []
    previous: tuple[int, ...] = ()
    for slot in slots:
        path = slot.path + ((slot.data,) if with_data and slot.kind == "leaf" else ())
        shared = 0
        while shared < min(len(previous), len(path)) and previous[shared] == path[shared]:
            shared += 1
        repeated = min(shared, max(len(path) - 1, 0))
        entries = [Index(v=position) for position in path[repeated:]]
        if slot.kind == "grand":
            entries = [Index(v=0)]
        items.append(RowColItem(t=_ITEM_TYPE[slot.kind], r=repeated, i=slot.data, x=entries))
        previous = path
    return items


def render(
    shape: Shape, row_slots: list[Slot], column_slots: list[Slot]
) -> dict[tuple[int, int], CellOut]:
    """The cells of the table body and headers, keyed by (row, column) from the top left."""
    cells: dict[tuple[int, int], CellOut] = {}
    _render_headers(shape, column_slots, cells)
    groups = _group_records(shape)
    row_items = axis_items(row_slots, with_data=False)
    first_row = shape.header_rows
    first_column = len(shape.rows)
    for line, (slot, item) in enumerate(zip(row_slots, row_items, strict=True), first_row):
        _render_row_labels(shape, slot, item.r, line, cells)
        for offset, column in enumerate(column_slots):
            records = groups.get((slot.path, column.path))
            if records:
                cells[(line, first_column + offset)] = _aggregate(shape, column.data, records)
    return cells


def _render_headers(
    shape: Shape, columns: list[Slot], cells: dict[tuple[int, int], CellOut]
) -> None:
    rows, fields, values = len(shape.rows), shape.columns, shape.values
    last = shape.header_rows - 1
    for position, field in enumerate(shape.rows):
        cells[(last, position)] = (shape.source.columns[field].name, None)
    if not fields and len(values) == 1:
        cells[(0, rows)] = (values[0].caption, None)
        return
    if len(values) == 1:
        cells[(0, 0)] = (values[0].caption, None)
    for position, field in enumerate(fields):
        cells[(0, rows + position)] = (shape.source.columns[field].name, None)
    if len(values) > 1:
        cells[(0, rows + len(fields))] = ("Values", None)
    items = axis_items(columns, with_data=len(values) > 1)
    for offset, (slot, item) in enumerate(zip(columns, items, strict=True)):
        caption = values[slot.data].caption
        suffix = caption if len(values) > 1 else "Total"
        match slot.kind:
            case "grand":
                text = f"Total {caption}" if len(values) > 1 else "Grand Total"
                cells[(1, rows + offset)] = (text, None)
            case "subtotal":
                level = len(slot.path)
                label = _label(shape, fields, level - 1, slot.path)
                cells[(level, rows + offset)] = (f"{_text(label[0])} {suffix}", None)
            case "leaf":
                for level in range(item.r, len(fields)):
                    cells[(1 + level, rows + offset)] = _label(shape, fields, level, slot.path)
                if len(values) > 1:
                    cells[(1 + len(fields), rows + offset)] = (caption, None)


def _render_row_labels(
    shape: Shape, slot: Slot, repeated: int, line: int, cells: dict[tuple[int, int], CellOut]
) -> None:
    fields = shape.rows
    match slot.kind:
        case "grand":
            cells[(line, 0)] = ("Grand Total", None)
        case "subtotal":
            level = len(slot.path) - 1
            label = _label(shape, fields, level, slot.path)
            cells[(line, level)] = (f"{_text(label[0])} Total", None)
        case "leaf":
            for level in range(repeated, len(fields)):
                cells[(line, level)] = _label(shape, fields, level, slot.path)


def _label(shape: Shape, fields: list[int], level: int, path: tuple[int, ...]) -> CellOut:
    field = fields[level]
    value = shape.items[field].label(path[level])
    column = shape.source.columns[field]
    if value is None:
        return "(blank)", None
    return value, None if column.number_format == "General" else column.number_format


def _text(value: CellContent) -> str:
    match value:
        case dt.datetime() if value.time() == dt.time():
            return value.date().isoformat()
        case dt.datetime():
            return value.isoformat(sep=" ")
        case float() if value.is_integer():
            return str(int(value))
        case _:
            return str(value)


def _key(shape: Shape, fields: list[int], record: int) -> tuple[int, ...]:
    return tuple(shape.items[field].record_rank(record) for field in fields)


def _group_records(shape: Shape) -> dict[GroupKey, list[int]]:
    """Record numbers for every row group and column group, at every depth."""
    groups: dict[GroupKey, list[int]] = {}
    for record in range(shape.source.record_count):
        row_key = _key(shape, shape.rows, record)
        column_key = _key(shape, shape.columns, record)
        for row_depth in range(len(row_key) + 1):
            for column_depth in range(len(column_key) + 1):
                key = (row_key[:row_depth], column_key[:column_depth])
                groups.setdefault(key, []).append(record)
    return groups


def _aggregate(shape: Shape, data: int, records: list[int]) -> CellOut:
    spec = shape.values[data]
    column = shape.source.columns[spec.field]
    present: list[Value] = [
        column.values[record] for record in records if column.values[record] is not None
    ]
    if spec.function == "count":
        return (len(present) or None), None
    numbers = [value for value in present if isinstance(value, int | float)]
    if not numbers:
        return None, None
    fmt = None if column.number_format == "General" else column.number_format
    match spec.function:
        case "sum":
            return sum(numbers), fmt
        case "average":
            return sum(numbers) / len(numbers), fmt
        case "min":
            return min(numbers), fmt
        case _:
            return max(numbers), fmt
