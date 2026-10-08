"""The pivot cache: a snapshot of the source data that a PivotTable reads from."""

import datetime as dt

from openpyxl.pivot.cache import (
    CacheDefinition,
    CacheField,
    CacheSource,
    FieldGroup,
    GroupItems,
    RangePr,
    SharedItems,
    WorksheetSource,
)
from openpyxl.pivot.fields import DateTimeField, Index, Missing, Number, Text
from openpyxl.pivot.record import Record, RecordList
from openpyxl.utils.datetime import to_excel

from excel_mcp.operations.pivot_calc import CalcField
from excel_mcp.operations.pivot_fields import FieldSetup, Shared
from excel_mcp.operations.pivot_groups import Grouping
from excel_mcp.operations.pivot_source import Column, Value

CacheItem = Missing | Number | Text | DateTimeField


def build_cache(setup: FieldSetup, calculated: list[CalcField]) -> CacheDefinition:
    """Describe the source's fields, then the fields derived from them."""
    source = setup.source
    fields = [
        CacheField(
            name=column.name,
            numFmtId=column.number_format_id,
            sharedItems=_shared_items(column, setup.shared.get(position)),
            fieldGroup=_base_group(setup, position),
        )
        for position, column in enumerate(source.columns)
    ]
    for position, groups in setup.date_groups.items():
        for index, grouping in groups:
            name = f"{grouping.by.capitalize()} ({source.columns[position].name})"
            fields.append(_derived(index, name, position, grouping))
    fields += [
        CacheField(name=item.name, numFmtId=0, formula=item.formula, databaseField=False)
        for item in calculated
    ]
    cache = CacheDefinition(
        cacheSource=CacheSource(
            type="worksheet",
            worksheetSource=WorksheetSource(ref=source.ref, sheet=source.sheet),
        ),
        cacheFields=fields,
        refreshedBy="excel-mcp-server",
        refreshedDate=to_excel(dt.datetime.now(dt.UTC).replace(tzinfo=None)),
        createdVersion=6,
        refreshedVersion=6,
        minRefreshableVersion=3,
        recordCount=source.record_count,
    )
    cache.records = RecordList(r=_records(setup))
    return cache


def _base_group(setup: FieldSetup, position: int) -> FieldGroup | None:
    if position in setup.number_groups:
        grouping = setup.number_groups[position]
        explicit_start, explicit_end = grouping.explicit
        return FieldGroup(
            base=position,
            rangePr=RangePr(
                autoStart=False if explicit_start else None,
                autoEnd=False if explicit_end else None,
                groupBy="range",
                startNum=float(grouping.start),  # pyright: ignore[reportArgumentType]
                endNum=float(grouping.end),  # pyright: ignore[reportArgumentType]
                groupInterval=grouping.interval,
            ),
            groupItems=_group_items(grouping),
        )
    if position in setup.date_groups:
        return FieldGroup(par=setup.date_groups[position][-1][0])
    return None


def _derived(index: int, name: str, base: int, grouping: Grouping) -> CacheField:
    assert isinstance(grouping.start, dt.datetime) and isinstance(grouping.end, dt.datetime)
    return CacheField(
        name=name,
        numFmtId=0,
        databaseField=False,
        fieldGroup=FieldGroup(
            base=base,
            rangePr=RangePr(
                autoStart=None,
                autoEnd=None,
                groupBy=grouping.by,
                startDate=grouping.start,
                endDate=grouping.end,
                groupInterval=None,
            ),
            groupItems=_group_items(grouping),
        ),
    )


def _group_items(grouping: Grouping) -> GroupItems:
    return GroupItems(s=[Text(v=label) for label in grouping.labels])


def _shared_items(column: Column, shared: Shared | None) -> SharedItems:
    present = [value for value in column.values if value is not None]
    blank = len(present) < len(column.values)
    numbers = [float(value) for value in present if isinstance(value, int | float)]
    moments = [value for value in present if isinstance(value, dt.datetime)]
    is_number, is_date = column.kind == "number", column.kind == "date"
    return SharedItems(
        _fields=[_item(value) for value in shared.values] if shared else (),
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


def _records(setup: FieldSetup) -> list[Record]:
    source = setup.source
    return [
        Record(
            _fields=[
                Index(v=setup.shared[position].index[row])
                if position in setup.shared
                else _item(column.values[row])
                for position, column in enumerate(source.columns)
            ]
        )
        for row in range(source.record_count)
    ]
