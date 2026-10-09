"""What a slicer or timeline does to the PivotTables it is connected to.

A slicer names the items of a field to show. Excel keeps them in the field's item list
(``h`` marks the hidden ones), which needs the field's items in the pivot cache. The figures in
the cells follow, so a PivotTable that hides items is made again, see `pivot_rebuild`.
"""

import datetime as dt
from dataclasses import dataclass

from openpyxl.pivot.cache import CacheDefinition
from openpyxl.pivot.fields import Index
from openpyxl.pivot.table import FieldItem, TableDefinition
from openpyxl.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from excel_mcp.errors import InvalidArgumentError
from excel_mcp.operations.pivot_definition import date_filter
from excel_mcp.operations.pivot_fields import distinct, item_text
from excel_mcp.operations.pivot_index import sheet_pivots
from excel_mcp.operations.pivot_options import DatePeriod
from excel_mcp.operations.pivot_rebuild import can_rebuild, field_labels, rebuild
from excel_mcp.operations.pivot_source import Value
from excel_mcp.text import quoted
from excel_mcp.workspace import worksheets


@dataclass(frozen=True)
class Connected:
    """A PivotTable and the sheet it is on."""

    sheet: Worksheet
    pivot: TableDefinition


def source_field(cache: CacheDefinition, name: str) -> int:
    """The position of a source column among the cache's fields."""
    columns = [
        (position, field.name)
        for position, field in enumerate(cache.cacheFields)
        if field.formula is None and not (field.databaseField is False and field.fieldGroup)
    ]
    for position, column in columns:
        if column.strip().casefold() == name.strip().casefold():
            if cache.cacheFields[position].fieldGroup is not None:
                raise InvalidArgumentError(
                    f"Field {column!r} is grouped; slicers need an ungrouped field."
                )
            return position
    raise InvalidArgumentError(
        f"Field {name!r} not found. Fields: {quoted(column for _, column in columns)}."
    )


def values_of(cache: CacheDefinition, position: int) -> list[Value]:
    return field_labels(cache.cacheFields[position])


def ensure_items(cache: CacheDefinition, position: int) -> None:
    """Give the cache the field's items, as Excel does when a slicer needs them."""
    field = cache.cacheFields[position]
    if field.sharedItems is None:  # a date group keeps its items in its fieldGroup
        return
    if field.sharedItems._fields or cache.records is None:  # pyright: ignore[reportAttributeAccessIssue]
        return
    records = cache.records.r
    inline = [record._fields[position] for record in records]
    shared = distinct([getattr(item, "v", None) for item in inline])
    first_seen = {}
    for item, index in zip(inline, shared.index, strict=True):
        first_seen.setdefault(index, item)
    field.sharedItems._fields = [first_seen[i] for i in range(len(first_seen))]  # pyright: ignore[reportAttributeAccessIssue]
    for record, index in zip(records, shared.index, strict=True):
        record._fields[position] = Index(v=index)


def item_order(cache: CacheDefinition, position: int) -> list[int]:
    """The field's items from the first to show to the last, as positions among its items."""
    values = values_of(cache, position)
    return distinct(values).order if values else []


def ensure_pivot_items(pivot: TableDefinition, position: int) -> None:
    field = pivot.pivotFields[position]
    if not field.items:
        order = item_order(pivot.cache, position)
        field.items = [FieldItem(x=x) for x in order] + [FieldItem(t="default")]


def shown_items(pivot: TableDefinition, position: int) -> list[int] | None:
    """The items of a field the PivotTable shows, or None when it has none listed."""
    items = pivot.pivotFields[position].items
    return [i.x for i in items if i.t != "default" and not i.h] if items else None


def choose_items(cache: CacheDefinition, position: int, names: list[str]) -> set[int]:
    """The items (positions among the cache's) whose text is in ``names``."""
    texts = {
        item_text(value).casefold(): index for index, value in enumerate(values_of(cache, position))
    }
    missing = [name for name in names if name.strip().casefold() not in texts]
    if missing:
        values = values_of(cache, position)
        listed = [item_text(values[x]) for x in item_order(cache, position)][:20]
        raise InvalidArgumentError(
            f"Field {cache.cacheFields[position].name!r} has no item {missing[0]!r}. "
            f"Items: {listed}."
        )
    return {texts[name.strip().casefold()] for name in names}


def prepare(group: list[Connected], position: int) -> None:
    """Make the items of a field, and of every field these PivotTables list, available."""
    cache = group[0].pivot.cache
    listed = {position}
    for connected in group:
        listed |= {p for p, f in enumerate(connected.pivot.pivotFields) if f.items}
    for field in sorted(listed):
        ensure_items(cache, field)
        for connected in group:
            ensure_pivot_items(connected.pivot, field)


def show_only(
    workbook: Workbook, group: list[Connected], position: int, chosen: set[int], max_cells: int
) -> bool:
    """Limit the field to the chosen items in every PivotTable of the group.

    Returns whether the figures in the cells were recalculated; see `_apply`."""
    prepare(group, position)
    changed = False
    for connected in group:
        for item in connected.pivot.pivotFields[position].items:
            if item.t != "default":
                hide = item.x not in chosen
                changed |= bool(item.h) != hide
                item.h = True if hide else None
    return changed and _apply(workbook, group, lambda request: request, max_cells)


def limit_dates(
    workbook: Workbook,
    group: list[Connected],
    position: int,
    period: tuple[dt.datetime, dt.datetime] | None,
    max_cells: int,
) -> bool:
    """Keep the records dated within the period in every PivotTable of the group.

    Returns whether the figures in the cells were recalculated; see `_apply`."""
    prepare(group, position)
    name = group[0].pivot.cache.cacheFields[position].name
    if not all(can_rebuild(c.pivot) for c in group):
        _set_period(group, position, name, period)
        return _apply(workbook, group, lambda request: request, max_cells)

    def change(request):
        others = [p for p in request.periods if p.field != name]
        kept = [DatePeriod(name, *period)] if period else []
        return type(request)(**{**vars(request), "periods": [*others, *kept]})

    return _apply(workbook, group, change, max_cells)


def _set_period(
    group: list[Connected], position: int, name: str, period: tuple[dt.datetime, dt.datetime] | None
) -> None:
    """Put the date filter on PivotTables that cannot be made again, as Excel stores it."""
    for connected in group:
        pivot = connected.pivot
        pivot.filters = [
            f for f in pivot.filters if not (f.fld == position and f.type == "dateBetween")
        ]
        if period:
            number = max((f.id or 0 for f in pivot.filters), default=0) + 1
            pivot.filters.append(date_filter(position, DatePeriod(name, *period), number))


def _apply(workbook, group, change, max_cells) -> bool:
    """Recalculate the cells of the PivotTables, or for PivotTables this server did not make,
    leave them and ask Excel to refresh when the file is opened. True when recalculated."""
    if not all(can_rebuild(c.pivot) for c in group):
        group[0].pivot.cache.refreshOnLoad = True
        return False
    _rebuild_all(workbook, group, change, max_cells)
    return True


def _rebuild_all(workbook, group, change, max_cells) -> None:
    made = [rebuild(workbook, c.sheet, c.pivot, change, max_cells) for c in group]
    for connected, pivot in zip(group, made, strict=True):
        pivot.cache = made[0].cache
        connected.pivot.cache = made[0].cache
    group[:] = [Connected(c.sheet, p) for c, p in zip(group, made, strict=True)]


def data_dates(cache: CacheDefinition, position: int) -> tuple[dt.datetime, dt.datetime]:
    shared = cache.cacheFields[position].sharedItems
    if shared.minDate is None or shared.maxDate is None:
        raise InvalidArgumentError(f"Field {cache.cacheFields[position].name!r} holds no dates.")
    return shared.minDate, shared.maxDate


def share_caches(workbook: Workbook) -> None:
    """Make PivotTables with the same cache id use one cache object, as the file has one cache.

    openpyxl reads a cache again for each sheet."""
    first: dict[int, CacheDefinition] = {}
    for sheet in worksheets(workbook):
        for pivot in sheet_pivots(sheet):
            pivot.cache = first.setdefault(pivot.cacheId, pivot.cache)
