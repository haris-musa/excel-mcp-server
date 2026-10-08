"""Adding slicers and timelines, as Insert > Slicer and Insert > Timeline do."""

import datetime as dt
import re
from collections.abc import Callable
from dataclasses import dataclass
from typing import Literal

from openpyxl.pivot.table import TableDefinition
from openpyxl.workbook import Workbook
from openpyxl.worksheet.table import Table
from openpyxl.worksheet.worksheet import Worksheet
from pydantic import Field

from excel_mcp.errors import InvalidArgumentError
from excel_mcp.inputs import InputModel
from excel_mcp.operations import slicer_pivot as pivots
from excel_mcp.operations import slicer_table
from excel_mcp.operations import slicer_xml as xml
from excel_mcp.operations.comparison import FormulaResults
from excel_mcp.operations.filters import table_area
from excel_mcp.operations.slicer_index import (
    caches,
    entries,
    find_source,
    sheet_number,
    table_number,
)
from excel_mcp.operations.slicer_place import add_cache, add_entry, anchor, next_shape_id
from excel_mcp.package import state_of
from excel_mcp.refs import CellRange, parse_cell

EMU_PER_CM = 360_000
_SLICER_SIZE = (1_828_800, 2_667_000)
_TIMELINE_SIZE = (3_333_750, 1_371_600)
Results = Callable[[Worksheet, CellRange], FormulaResults]


class Source(InputModel):
    sheet: str = Field(description="Sheet the table or PivotTable is on.")
    name: str = Field(description="Table or PivotTable name, as describe_sheet lists it.")


class Timeline(InputModel):
    level: Literal["years", "quarters", "months", "days"] = Field(
        default="months", description="Unit of the time scale."
    )
    start: dt.date | None = Field(
        default=None, description="First day of the period shown, e.g. '2025-03-01'. With end."
    )
    end: dt.date | None = Field(default=None, description="Last day of the period shown.")


@dataclass(frozen=True)
class SlicerRequest:
    source: Source
    connect: list[Source]
    field: str
    cell: str
    width_cm: float | None
    height_cm: float | None
    caption: str | None
    name: str | None
    columns: int
    style: str | None
    selected: list[str]
    sort: Literal["ascending", "descending"]
    hide_empty: bool
    timeline: Timeline | None = None


def add_slicer(
    workbook: Workbook, sheet: Worksheet, request: SlicerRequest, results: Results, max_cells: int
) -> str:
    """Add the slicer or timeline to ``sheet``; returns its name."""
    owner, target = find_source(workbook, request.source.sheet, request.source.name)
    if isinstance(target, Table):
        if request.timeline or request.connect:
            raise InvalidArgumentError(
                "A table slicer connects to its table only, and tables have no timelines."
            )
        return _table_slicer(workbook, sheet, owner, target, request, results)
    pivots.share_caches(workbook)
    group = _connected(workbook, owner, target, request)
    if request.timeline:
        return _timeline(workbook, sheet, group, request, max_cells)
    return _pivot_slicer(workbook, sheet, group, request, max_cells)


def _table_slicer(
    workbook: Workbook,
    sheet: Worksheet,
    owner: Worksheet,
    table: Table,
    request: SlicerRequest,
    results: Results,
) -> str:
    _check_style(request.style, "slicer")
    _, column = slicer_table.table_column(table, request.field)
    _check_free_column(workbook, table, column.name)
    name, cache = _names(workbook, request, "Slicer_")
    if request.selected:
        slicer_table.show_only(
            owner, table, column.name, request.selected, results(owner, table_area(table))
        )
    options = xml.CacheOptions(request.sort, request.hide_empty)
    number = table_number(state_of(workbook), table)
    text = xml.table_cache(cache, column.name, number, column.id, options)
    entry = xml.slicer_entry(
        name, cache, request.caption or column.name, request.columns, True, request.style
    )
    _add(workbook, sheet, "table", request, (name, cache), (text, entry))
    return name


def _pivot_slicer(
    workbook: Workbook,
    sheet: Worksheet,
    group: list[pivots.Connected],
    request: SlicerRequest,
    max_cells: int,
) -> str:
    _check_style(request.style, "slicer")
    cache = group[0].pivot.cache
    position = pivots.source_field(cache, request.field)
    field = cache.cacheFields[position].name
    _check_free(workbook, group, field, "pivot")
    if request.selected:
        pivots.ensure_items(cache, position)
        chosen = pivots.choose_items(cache, position, request.selected)
        pivots.show_only(workbook, group, position, chosen, max_cells)
    else:
        pivots.prepare(group, position)
    name, cache_name = _names(workbook, request, "Slicer_")
    first = group[0].pivot
    options = xml.CacheOptions(request.sort, request.hide_empty)
    text = xml.pivot_cache(
        cache_name,
        field,
        _tabs(workbook, group),
        _pivot_cache_id(workbook, first),
        _items(first, position, request.sort),
        options,
    )
    entry = xml.slicer_entry(
        name, cache_name, request.caption or field, request.columns, True, request.style
    )
    _add(workbook, sheet, "pivot", request, (name, cache_name), (text, entry))
    return name


def _timeline(
    workbook: Workbook,
    sheet: Worksheet,
    group: list[pivots.Connected],
    request: SlicerRequest,
    max_cells: int,
) -> str:
    timeline = request.timeline
    assert timeline is not None
    _check_style(request.style, "timeline")
    if (
        request.selected
        or request.columns != 1
        or request.hide_empty
        or request.sort != "ascending"
    ):
        raise InvalidArgumentError(
            "A timeline takes no selected_items, columns, sort or hide_empty_items; give "
            "timeline.start and timeline.end to select a period."
        )
    if (timeline.start is None) != (timeline.end is None):
        raise InvalidArgumentError("Give both timeline.start and timeline.end, or neither.")
    cache = group[0].pivot.cache
    position = pivots.source_field(cache, request.field)
    field = cache.cacheFields[position].name
    _check_free(workbook, group, field, "timeline")
    pivots.prepare(group, position)
    first, last = pivots.data_dates(cache, position)
    bounds = xml.timeline_bounds(first.date(), last.date())
    selection = None
    if timeline.start is not None and timeline.end is not None:
        if timeline.start > timeline.end:
            raise InvalidArgumentError("timeline.start is after timeline.end.")
        selection = (timeline.start, timeline.end)
        period = tuple(dt.datetime.combine(day, dt.time()) for day in selection)
        pivots.limit_dates(workbook, group, position, period, max_cells)  # pyright: ignore[reportArgumentType]
    name, cache_name = _names(workbook, request, "NativeTimeline_")
    text = xml.timeline_cache(
        cache_name,
        field,
        _tabs(workbook, group),
        _pivot_cache_id(workbook, group[0].pivot),
        bounds,
        selection,
    )
    entry = xml.timeline_entry(
        name,
        cache_name,
        request.caption or field,
        xml.LEVELS.index(timeline.level),
        xml.timeline_scroll(bounds),
        request.style,
    )
    _add(workbook, sheet, "timeline", request, (name, cache_name), (text, entry))
    return name


def _connected(
    workbook: Workbook, owner: Worksheet, pivot: TableDefinition, request: SlicerRequest
) -> list[pivots.Connected]:
    group = [pivots.Connected(owner, pivot)]
    for other in request.connect:
        sheet, found = find_source(workbook, other.sheet, other.name)
        if not isinstance(found, TableDefinition) or found.cacheId != pivot.cacheId:
            raise InvalidArgumentError(
                f"{other.name!r} is not a PivotTable that shares {pivot.name!r}'s data cache "
                "(as copies of a sheet do), so one slicer cannot filter both."
            )
        if found is pivot:
            raise InvalidArgumentError(f"{other.name!r} is the PivotTable the slicer is for.")
        group.append(pivots.Connected(sheet, found))
    return group


def _tabs(workbook: Workbook, group: list[pivots.Connected]) -> list[tuple[int, str]]:
    state = state_of(workbook)
    return [(sheet_number(state, c.sheet), c.pivot.name) for c in group]


def _names(workbook: Workbook, request: SlicerRequest, stem: str) -> tuple[str, str]:
    """The slicer's name and its cache's, unique in the workbook as Excel makes them."""
    taken = {entry.name.casefold() for entry in entries(workbook)}
    if request.name is not None:
        if request.name.casefold() in taken:
            raise InvalidArgumentError(f"A slicer or timeline named {request.name!r} exists.")
        name = request.name
    else:
        name, number = request.field, 0
        while name.casefold() in taken:
            number += 1
            name = f"{request.field} {number}"
    stem += re.sub(r"\W", "_", request.field)
    used = {n.casefold() for n in workbook.defined_names} | {c.casefold() for c in caches(workbook)}
    cache, number = stem, 0
    while cache.casefold() in used:
        number += 1
        cache = f"{stem}{number}"
    return name, cache


def _check_style(style: str | None, kind: str) -> None:
    wanted = "TimeSlicerStyle" if kind == "timeline" else "SlicerStyle"
    if style is not None and not re.fullmatch(wanted + r"(Light|Dark|Other)\d", style):
        raise InvalidArgumentError(
            f"style {style!r} must be a built-in {kind} style such as {wanted}Light1 "
            f"({wanted}Dark3, {wanted}Other1)."
        )


def _check_free(workbook: Workbook, group: list[pivots.Connected], field: str, kind: str) -> None:
    """A field has one slicer, and one timeline, per PivotTable."""
    mine = set(_tabs(workbook, group))
    for name, (_, info) in caches(workbook).items():
        if info.field == field and info.kind == kind and mine & set(info.pivots):
            raise InvalidArgumentError(
                f"Field {field!r} already has the {kind} {name!r} on these PivotTables; "
                "delete_slicer it first."
            )


def _check_free_column(workbook: Workbook, table: Table, column: str) -> None:
    number = table_number(state_of(workbook), table)
    for name, (_, info) in caches(workbook).items():
        if info.kind == "table" and info.table_id == number and info.field == column:
            raise InvalidArgumentError(
                f"Column {column!r} of table {table.displayName!r} already has the slicer "
                f"cache {name!r}; delete_slicer it first."
            )


def _pivot_cache_id(workbook: Workbook, pivot: TableDefinition) -> int:
    """The id slicer caches find the PivotTable's cache by; Excel keeps it in the cache."""
    extensions = state_of(workbook).workbook.cache_extensions
    found = re.search(r'pivotCacheId="(\d+)"', extensions.get(str(pivot.cacheId), ""))
    if found:
        return int(found[1])
    number = 500_000_000 + pivot.cacheId
    extensions[str(pivot.cacheId)] = xml.cache_id_extension(number)
    return number


def _items(pivot: TableDefinition, position: int, sort: str) -> list[tuple[int, bool]]:
    shown = pivots.shown_items(pivot, position) or []
    listed = [(x, x in shown) for x in pivots.item_order(pivot.cache, position)]
    return listed[::-1] if sort == "descending" else listed


def _add(
    workbook: Workbook,
    sheet: Worksheet,
    kind: str,
    request: SlicerRequest,
    names: tuple[str, str],
    texts: tuple[str, str],
) -> None:
    """Add the cache, the slicer's entry and its shape."""
    name, cache = names
    add_cache(workbook, kind, cache, texts[0])
    row, column = parse_cell(request.cell)
    default = _TIMELINE_SIZE if kind == "timeline" else _SLICER_SIZE
    size = (
        round(request.width_cm * EMU_PER_CM) if request.width_cm else default[0],
        round(request.height_cm * EMU_PER_CM) if request.height_cm else default[1],
    )
    package = state_of(workbook).sheet(sheet)
    shape = anchor(sheet, kind, name, (row, column), size, next_shape_id(package))
    add_entry(workbook, sheet, kind, texts[1], shape)
