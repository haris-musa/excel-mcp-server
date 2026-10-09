"""Listing and deleting the slicers and timelines of a sheet."""

import re
import uuid
from xml.sax.saxutils import quoteattr

from openpyxl.pivot.table import TableDefinition
from openpyxl.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet
from pydantic import BaseModel

from excel_mcp.errors import InvalidArgumentError
from excel_mcp.operations import drawings, slicer_table
from excel_mcp.operations import slicer_pivot as pivots
from excel_mcp.operations import slicer_xml as xml
from excel_mcp.operations.pivot_fields import item_text
from excel_mcp.operations.slicer_index import (
    Entry,
    caches,
    entries,
    pivot_of,
    sheet_number,
    sheet_of,
    table_number,
    table_of,
)
from excel_mcp.operations.slicer_place import add_cache, add_entry, next_shape_id
from excel_mcp.package import state_of
from excel_mcp.package.anchors import Corner
from excel_mcp.text import quoted


class SlicerInfo(BaseModel):
    name: str
    kind: str
    caption: str
    target: str
    field: str
    range: str | None = None
    selected_items: list[str] = []
    level: str | None = None
    start: str | None = None
    end: str | None = None
    style: str | None = None


def list_slicers(workbook: Workbook, sheet: Worksheet) -> list[SlicerInfo]:
    known = caches(workbook)
    found = []
    for entry in entries(workbook, sheet):
        info = known.get(entry.values["cache"])
        if info is None:
            continue
        found.append(_describe(workbook, entry, info[1]))
    return found


def _describe(workbook: Workbook, entry: Entry, cache: xml.CacheInfo) -> SlicerInfo:
    values = entry.values
    selected: list[str] = []
    target = ""
    if cache.kind == "table":
        owner, table = table_of(workbook, cache.table_id or 0)
        target = f"{owner.title}!{table.displayName}"
        index, _ = slicer_table.table_column(table, cache.field)
        selected = slicer_table.shown_values(table, index) or []
    else:
        owner, pivot = pivot_of(workbook, *cache.pivots[0])
        target = ", ".join(f"{pivot_of(workbook, *p)[0].title}!{p[1]}" for p in cache.pivots)
        if cache.kind == "pivot":
            selected = _selected_items(pivot, cache)
    chosen = cache.selection
    return SlicerInfo(
        name=values["name"],
        kind="timeline"
        if entry.kind == "timeline"
        else "table"
        if cache.kind == "table"
        else "pivot",
        caption=values.get("caption", ""),
        target=target,
        field=cache.field,
        range=_range(workbook, entry),
        selected_items=selected,
        level=xml.LEVELS[int(values.get("level", "2"))] if entry.kind == "timeline" else None,
        start=f"{chosen[0]:%Y-%m-%d}" if chosen else None,
        end=f"{chosen[1]:%Y-%m-%d}" if chosen else None,
        style=values.get("style"),
    )


def _selected_items(pivot: TableDefinition, cache: xml.CacheInfo) -> list[str]:
    """The items a slicer shows as selected; empty when all are."""
    if all(on for _, on in cache.items):
        return []
    names = [f.name for f in pivot.cache.cacheFields]
    values = pivots.values_of(pivot.cache, names.index(cache.field))
    chosen = [item_text(values[x]) for x, on in cache.items if on]
    return chosen[::-1] if cache.sort == "descending" else chosen


def _range(workbook: Workbook, entry: Entry) -> str | None:
    return placed_range(workbook, entry.sheet, entry.name)


def placed_range(workbook: Workbook, sheet: Worksheet, name: str) -> str | None:
    """The cells the shape of a slicer or timeline covers."""
    shape = _anchor_of(state_of(workbook).sheet(sheet).anchors, name)
    return drawings.anchored_range(shape) if shape else None


def _anchor_of(anchors: list[str], name: str) -> str | None:
    wanted = f'name="{name}"'
    return next(
        (a for a in anchors if re.search(rf"<\w+:(?:slicer|timeslicer)\b[^>]*{wanted}", a)), None
    )


def delete_slicer(workbook: Workbook, sheet: Worksheet, name: str, max_cells: int) -> str:
    """Remove a slicer or timeline. Returns what it was connected to."""
    found = [e for e in entries(workbook, sheet) if e.name.casefold() == name.strip().casefold()]
    if not found:
        available = [e.name for e in entries(workbook, sheet)]
        raise InvalidArgumentError(
            f"Sheet {sheet.title!r} has no slicer or timeline {name!r}. "
            f"Available: {quoted(available) if available else 'none'}."
        )
    pivots.share_caches(workbook)
    entry = found[0]
    info = caches(workbook)[entry.values["cache"]][1]
    _release(workbook, entry, info, max_cells)
    remove_entry(workbook, entry)
    return info.field


def remove_entry(workbook: Workbook, entry: Entry) -> None:
    """Take a slicer or timeline out of its sheet; its cache goes when nothing uses it."""
    package = state_of(workbook).sheet(entry.sheet)
    text = entry.part.data.decode("utf-8")
    text = re.sub(rf'<(?:slicer|timeline)\b[^>]*\bname="{re.escape(entry.name)}"[^>]*/>', "", text)
    entry.part.data = text.encode("utf-8")
    if not xml.read_entries(text):
        package.links.remove(entry.link)
        for uri, extension in list(package.extensions.items()):
            if entry.link.id in re.findall(r'r:id="([^"]*)"', extension):
                del package.extensions[uri]
    shape = _anchor_of(package.anchors, entry.name)
    if shape is not None:
        package.anchors.remove(shape)


def drop_table_slicers(
    workbook: Workbook, sheet: Worksheet, dead: dict[str, frozenset[str] | None]
) -> None:
    """Remove the slicers of tables, and of table columns, that an edit deletes, as Excel does."""
    known = caches(workbook)
    for entry in entries(workbook):
        info = known[entry.values["cache"]][1]
        if info.kind != "table":
            continue
        owner, table = table_of(workbook, info.table_id or 0)
        gone = dead.get(table.displayName.casefold(), frozenset())
        if owner is sheet and (gone is None or info.field.casefold() in gone):
            remove_entry(workbook, entry)


def _release(workbook: Workbook, entry: Entry, info: xml.CacheInfo, max_cells: int) -> None:
    """Take away the filtering only the slicer kept: a timeline's dates, a field's hidden items
    that no axis shows. Excel keeps the rest, as a filter of the table or PivotTable."""
    if info.kind == "table":
        return
    group = [pivots.Connected(*pivot_of(workbook, *p)) for p in info.pivots]
    cache = group[0].pivot.cache
    position = [f.name for f in cache.cacheFields].index(info.field)
    if info.kind == "timeline":
        if info.selection:
            pivots.limit_dates(workbook, group, position, None, max_cells)
    elif not _on_axis(group, position):
        every = set(range(len(pivots.values_of(cache, position))))
        pivots.show_only(workbook, group, position, every, max_cells)


def _on_axis(group: list[pivots.Connected], position: int) -> bool:
    return any(
        position in {f.x for f in [*c.pivot.rowFields, *c.pivot.colFields]}
        or position in {f.fld for f in c.pivot.pageFields}
        for c in group
    )


__all__ = ["Corner", "SlicerInfo", "delete_slicer", "list_slicers", "sheet_of"]


# -- copying a sheet ---------------------------------------------------------------------------

_SHAPE_ID = re.compile(r'(<xdr:cNvPr\b[^>]*?\bid=")\d+(")')
_GUID = re.compile(r'(<a16:creationId\b[^>]*?\bid=")[^"]*(")')
_UID = re.compile(r'(\b\w+:uid=")[^"]*(")')


def copy_slicers(
    workbook: Workbook, source: Worksheet, target: Worksheet, tables: dict[str, str]
) -> None:
    """Give a copy of a sheet the slicers and timelines of the original, as Excel does.

    A cache whose PivotTables or table are all on the sheet is copied with its slicers, so the
    copy filters the copy. Otherwise the copies of the PivotTables join the cache, and slicers
    on the sheet get copies that use the same cache.
    """
    state = state_of(workbook)
    mine = entries(workbook, source)
    for name, (_, info) in list(caches(workbook).items()):
        users = [e for e in mine if e.values["cache"] == name]
        if info.kind == "table":
            owner, table = table_of(workbook, info.table_id or 0)
            copy = target.tables.get(tables.get(table.displayName.casefold(), ""))
            if owner is source and copy is not None:
                _copy_cache(workbook, name, users, target, table_id=table_number(state, copy))
            continue
        on_source = [p for p in info.pivots if sheet_of(state, p[0]) is source]
        if on_source and len(on_source) == len(info.pivots) and users:
            _copy_cache(workbook, name, users, target, tab_id=sheet_number(state, target))
            continue
        if on_source:
            _join_cache(workbook, name, on_source, sheet_number(state, target))
        for entry in users:
            _copy_entry(workbook, entry, target, entry.values["cache"], info.kind)


def _join_cache(workbook: Workbook, name: str, pivots: list[tuple[int, str]], tab: int) -> None:
    link = caches(workbook)[name][0]
    text = link.target.data.decode("utf-8")  # pyright: ignore[reportAttributeAccessIssue]
    added = "".join(f'<pivotTable tabId="{tab}" name="{p}"/>' for _, p in pivots)
    link.target.data = text.replace("</pivotTables>", added + "</pivotTables>").encode("utf-8")  # pyright: ignore[reportAttributeAccessIssue]


def _copy_cache(
    workbook: Workbook,
    name: str,
    users: list[Entry],
    target: Worksheet,
    *,
    tab_id: int | None = None,
    table_id: int | None = None,
) -> None:
    link, info = caches(workbook)[name]
    stem = re.sub(r"\d+$", "", name)
    taken = {n.casefold() for n in workbook.defined_names} | {
        n.casefold() for n in caches(workbook)
    }
    number = 1
    while f"{stem}{number}".casefold() in taken:
        number += 1
    new = f"{stem}{number}"
    text = link.target.data.decode("utf-8")  # pyright: ignore[reportAttributeAccessIssue]
    text = text.replace(f'name="{name}"', f'name="{new}"', 1)
    text = _UID.sub(lambda m: m[1] + _new_guid() + m[2], text, count=1)
    if tab_id is not None:
        text = re.sub(r'(<pivotTable\b[^>]*?\btabId=")\d+', rf"\g<1>{tab_id}", text)
    if table_id is not None:
        text = re.sub(r'(<x15:tableSlicerCache\b[^>]*?\btableId=")\d+', rf"\g<1>{table_id}", text)
    add_cache(workbook, info.kind, new, text)
    for entry in users:
        _copy_entry(workbook, entry, target, new, info.kind)


def _copy_entry(workbook: Workbook, entry: Entry, target: Worksheet, cache: str, kind: str) -> None:
    taken = {e.name.casefold() for e in entries(workbook)}
    stem = re.sub(r" \d+$", "", entry.name)
    number = 1
    while f"{stem} {number}".casefold() in taken:
        number += 1
    name = f"{stem} {number}"
    values = {k: _new_guid() if k.endswith(":uid") else v for k, v in entry.values.items()} | {
        "name": name,
        "cache": cache,
    }
    tag = entry.kind + "".join(f" {k}={quoteattr(v)}" for k, v in values.items()) + "/>"
    package = state_of(workbook).sheet(entry.sheet)
    shape = _anchor_of(package.anchors, entry.name)
    assert shape is not None
    shape = shape.replace(f'name="{entry.name}"', f'name="{name}"')
    shape = _SHAPE_ID.sub(
        rf"\g<1>{next_shape_id(state_of(workbook).sheet(target))}\g<2>", shape, count=1
    )
    shape = _GUID.sub(rf"\g<1>{_new_guid()}\g<2>", shape)
    add_entry(workbook, target, kind, "<" + tag, shape, entry.part.data.decode("utf-8"))


def _new_guid() -> str:
    return "{" + str(uuid.uuid4()).upper() + "}"
