"""Changing a PivotTable this server made by creating it again with one thing different.

The request that made the PivotTable is read back from its definition and its cache, changed,
and the PivotTable is made again where it was. The figures come from the source as it is now.
"""

import datetime as dt
from collections.abc import Callable

from openpyxl.pivot.cache import CacheDefinition, CacheField, RangePr
from openpyxl.pivot.table import AutoSortScope, DataField, PivotFilter, TableDefinition
from openpyxl.pivot.table import PivotField as PivotFieldDefinition
from openpyxl.styles.numbers import BUILTIN_FORMATS, BUILTIN_FORMATS_MAX_SIZE
from openpyxl.utils.datetime import from_excel
from openpyxl.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from excel_mcp.errors import InvalidArgumentError
from excel_mcp.operations.cells import clear_range
from excel_mcp.operations.pivot import PivotRequest, create_pivot
from excel_mcp.operations.pivot_cache import REFRESHED_BY
from excel_mcp.operations.pivot_definition import (
    PREVIOUS_ITEM,
    SHOW_AS_EXTENSION,
    SHOW_VALUES_AS,
    VALUES_FIELD,
)
from excel_mcp.operations.pivot_fields import item_text
from excel_mcp.operations.pivot_index import pivot_area, sheet_pivots
from excel_mcp.operations.pivot_options import (
    CalculatedField,
    DatePeriod,
    NumberGroup,
    PivotField,
    PivotValue,
    ShowAs,
)
from excel_mcp.operations.pivot_source import Value
from excel_mcp.operations.pivot_specs import COMPARES_WITH_ITEM, NEEDS_BASE_FIELD
from excel_mcp.package import state_of
from excel_mcp.refs import cell_name

_SHOW_AS_ATTRIBUTE: dict[str, ShowAs] = {
    "percentOfTotal": "percent_of_total",
    "percentOfRow": "percent_of_row",
    "percentOfCol": "percent_of_column",
    "difference": "difference_from",
    "percentDiff": "percent_difference_from",
    "percent": "percent_of",
    "runTotal": "running_total",
}
_SHOW_AS_EXTENSION: dict[str, ShowAs] = {value: key for key, value in SHOW_AS_EXTENSION.items()}


def can_rebuild(pivot: TableDefinition) -> bool:
    return pivot.cache.refreshedBy == REFRESHED_BY


def rebuild(
    workbook: Workbook,
    sheet: Worksheet,
    pivot: TableDefinition,
    change: Callable[[PivotRequest], PivotRequest],
    max_cells: int,
) -> TableDefinition:
    """Make the PivotTable again with ``change`` applied to the request that made it."""
    if not can_rebuild(pivot):
        raise InvalidArgumentError(
            f"PivotTable {pivot.name!r} was not made by create_pivot_table (or Excel refreshed "
            "it since), so its figures cannot be recalculated here. Change it in Excel."
        )
    request, source = request_of(workbook, sheet, pivot)
    area = pivot_area(pivot)
    pivots = sheet_pivots(sheet)
    position = pivots.index(pivot)
    pivots.remove(pivot)
    clear_range(sheet, str(area), contents=True, formats=False, max_cells=max_cells)
    source_sheet, reference = source
    create_pivot(
        workbook,
        workbook[source_sheet],
        reference,
        sheet,
        cell_name(area.min_row, area.min_col),
        change(request),
        max_cells,
        copy_of_name=True,
    )
    made = pivots.pop()
    made.cacheId = pivot.cacheId
    pivots.insert(position, made)
    return made


def request_of(
    workbook: Workbook, sheet: Worksheet, pivot: TableDefinition
) -> tuple[PivotRequest, tuple[str, str]]:
    """The request that makes ``pivot``, and its source as (sheet, range)."""
    cache = pivot.cache
    names = _Names(cache)
    source = cache.cacheSource.worksheetSource
    rows = _names(names, [f.x for f in pivot.rowFields or [] if f.x != VALUES_FIELD])
    columns = _names(names, [f.x for f in pivot.colFields or [] if f.x != VALUES_FIELD])
    filters = _names(names, [page.fld for page in pivot.pageFields or []])
    placed = [f.x for f in [*pivot.rowFields, *pivot.colFields] if f.x != VALUES_FIELD]
    on_axes = {names.base(x) for x in [*placed, *(page.fld for page in pivot.pageFields)]}
    fields: list[PivotField] = []
    filtered: list[PivotField] = []
    for position, field in enumerate(pivot.pivotFields):
        if position >= names.source_count or not field.items:
            continue
        option = _field_option(names, cache, pivot, position)
        if position in on_axes:
            fields.append(option)
        else:
            filtered.append(option)
    request = PivotRequest(
        rows=rows,
        columns=columns,
        values=[_value(workbook, sheet, pivot, names, i) for i in range(len(pivot.dataFields))],
        filters=filters,
        fields=[f for f in fields if f != PivotField(field=f.field)],
        calculated=[
            CalculatedField(name=f.name, formula=f.formula)
            for f in cache.cacheFields
            if f.formula is not None
        ],
        layout="compact" if pivot.compact else "outline" if pivot.outline else "tabular",
        subtotals=not any(
            pivot.pivotFields[names.base(f.x)].defaultSubtotal is False
            for f in [*pivot.rowFields, *pivot.colFields]
            if f.x != VALUES_FIELD
        ),
        values_in="rows" if pivot.dataOnRows else "columns",
        name=pivot.name,
        filtered=filtered,
        periods=[_period(names, f) for f in pivot.filters],
    )
    return request, (source.sheet, source.ref)


class _Names:
    """The fields of a cache: the source's columns, the groups made from them, calculations."""

    def __init__(self, cache: CacheDefinition) -> None:
        self.fields = cache.cacheFields
        self.source_count = sum(1 for f in self.fields if f.formula is None and not _is_group(f))

    def base(self, position: int) -> int:
        group = self.fields[position].fieldGroup
        return group.base if _is_group(self.fields[position]) and group else position

    def name(self, position: int) -> str:
        return self.fields[self.base(position)].name


def _is_group(field: CacheField) -> bool:
    return field.databaseField is False and field.fieldGroup is not None


def _names(names: _Names, positions: list[int]) -> list[str]:
    found: list[str] = []
    for position in positions:
        name = names.name(position)
        if name not in found:
            found.append(name)
    return found


def _field_option(
    names: _Names, cache: CacheDefinition, pivot: TableDefinition, position: int
) -> PivotField:
    field = pivot.pivotFields[position]
    group = cache.cacheFields[position].fieldGroup
    labels = field_labels(cache.cacheFields[position])
    items = [item for item in field.items if item.t != "default"]
    hidden = any(item.h for item in items)
    ranges = group.rangePr if group and group.rangePr else None
    derived = [
        f.fieldGroup.rangePr.groupBy
        for f in cache.cacheFields
        if _is_group(f) and f.fieldGroup.base == position and f.fieldGroup.rangePr
    ]
    scope = field.autoSortScope
    grouped = group is not None
    shown = [position] if group is None or not derived else _derived_positions(cache, position)
    return PivotField(
        field=names.name(position),
        show_items=[item_text(labels[item.x]) for item in items if not item.h] if hidden else [],
        sort="descending"
        if any(_descending(pivot.pivotFields[p], grouped) for p in shown)
        else "ascending",
        sort_by=_sort_by(pivot, scope) if scope else None,
        group_dates=list(reversed(derived)),
        group_numbers=_number_group(ranges) if ranges and ranges.groupBy == "range" else None,
    )


def _derived_positions(cache: CacheDefinition, base: int) -> list[int]:
    return [
        position
        for position, f in enumerate(cache.cacheFields)
        if _is_group(f) and f.fieldGroup.base == base
    ]


def _descending(field: PivotFieldDefinition, grouped: bool) -> bool:
    """Ranges and days are listed in the order given, with sortType manual."""
    if field.sortType != "manual" or not grouped:
        return field.sortType == "descending"
    items = [item.x for item in field.items if item.t != "default"]
    return len(items) > 2 and items[0] != 0 and items[-1] == 0


def field_labels(field: CacheField) -> list[Value]:
    """The items of a field as the values they stand for."""
    if field.fieldGroup is not None and field.fieldGroup.groupItems is not None:
        return [item.v for item in field.fieldGroup.groupItems.s]
    return [getattr(item, "v", None) for item in field.sharedItems._fields]  # pyright: ignore[reportAttributeAccessIssue,reportOptionalMemberAccess]


def _number_group(ranges: RangePr) -> NumberGroup:
    return NumberGroup(
        by=ranges.groupInterval or 1,
        start=None if ranges.autoStart is not False else ranges.startNum,
        end=None if ranges.autoEnd is not False else ranges.endNum,
    )


def _sort_by(pivot: TableDefinition, scope: AutoSortScope) -> str | None:
    reference = scope.pivotArea.references[0]
    return pivot.dataFields[reference.x[0].v].name


def _value(
    workbook: Workbook, sheet: Worksheet, pivot: TableDefinition, names: _Names, position: int
) -> PivotValue:
    data = pivot.dataFields[position]
    show_as = _show_as(workbook, sheet, pivot, data, position)
    base_field = names.fields[data.baseField].name if show_as in NEEDS_BASE_FIELD else None
    return PivotValue(
        field=names.fields[data.fld].name,
        function=data.subtotal,
        number_format=_number_format(workbook, data.numFmtId),
        show_as=show_as,
        base_field=base_field,
        base_item=_base_item(pivot, data) if show_as in COMPARES_WITH_ITEM else None,
    )


def _show_as(
    workbook: Workbook, sheet: Worksheet, pivot: TableDefinition, data: DataField, position: int
) -> ShowAs | None:
    if data.showDataAs != "normal":
        return _SHOW_AS_ATTRIBUTE[data.showDataAs]
    extension = (
        state_of(workbook).sheet(sheet).pivot_fields.get((pivot.name, "dataField", position), "")
    )
    if SHOW_VALUES_AS not in extension:
        return None
    return next(
        (value for key, value in _SHOW_AS_EXTENSION.items() if f'pivotShowAs="{key}"' in extension),
        None,
    )


def _base_item(pivot: TableDefinition, data: DataField) -> str | None:
    if data.baseItem == PREVIOUS_ITEM:
        return None
    field = pivot.pivotFields[data.baseField]
    item = [i for i in field.items if i.t != "default"][data.baseItem]
    return item_text(field_labels(pivot.cache.cacheFields[data.baseField])[item.x])


def _number_format(workbook: Workbook, number_format_id: int | None) -> str | None:
    if number_format_id is None:
        return None
    if number_format_id < BUILTIN_FORMATS_MAX_SIZE:
        return BUILTIN_FORMATS[number_format_id]
    return workbook._number_formats[number_format_id - BUILTIN_FORMATS_MAX_SIZE]  # pyright: ignore[reportAttributeAccessIssue]


def _period(names: _Names, pivot_filter: PivotFilter) -> DatePeriod:
    return DatePeriod(
        names.name(pivot_filter.fld),
        _date(pivot_filter.stringValue1 or ""),
        _date(pivot_filter.stringValue2 or ""),
    )


def _date(serial: str) -> dt.datetime:
    moment = from_excel(int(serial))
    if not isinstance(moment, dt.datetime):
        raise InvalidArgumentError(f"{serial!r} is not a date.")
    return moment
