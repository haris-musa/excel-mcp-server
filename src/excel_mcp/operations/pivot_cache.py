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
from openpyxl.pivot.fields import Boolean, DateTimeField, Error, Missing, Number, Text
from openpyxl.utils.datetime import to_excel

from excel_mcp.calc.values import ExcelError
from excel_mcp.errors import InvalidArgumentError
from excel_mcp.operations.pivot_calc import CalcField
from excel_mcp.operations.pivot_fields import FieldSetup, Shared
from excel_mcp.operations.pivot_groups import Grouping
from excel_mcp.operations.pivot_records import PivotRecords, SharedIndex
from excel_mcp.operations.pivot_source import Column, Value, is_number

REFRESHED_BY = "excel-mcp-server"
CacheItem = Missing | Number | Text | Boolean | Error | DateTimeField


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
        refreshedBy=REFRESHED_BY,
        refreshedDate=to_excel(dt.datetime.now(dt.UTC).replace(tzinfo=None)),
        createdVersion=6,
        refreshedVersion=6,
        minRefreshableVersion=3,
        recordCount=source.record_count,
    )
    cache.records = _records(setup)
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
    """The flags Excel writes: only those that differ from the defaults (a field with text,
    no blank and no dates has none), and no number range for a field that holds dates."""
    kinds = column.kinds
    blank = None in column.values
    ranged = "number" in kinds and "date" not in kinds
    numbers = [float(value) for value in column.values if is_number(value)]
    moments = [value for value in column.values if isinstance(value, dt.datetime)]
    return SharedItems(
        _fields=[cache_item(value) for value in shared.values] if shared else (),
        containsBlank=True if blank else None,
        containsSemiMixedTypes=False if "other" not in kinds and not blank else None,
        containsString=False if "other" not in kinds else None,
        containsMixedTypes=True if len(kinds) > 1 else None,
        containsNumber=True if ranged else None,
        containsInteger=True if ranged and all(n.is_integer() for n in numbers) else None,
        minValue=min(numbers) if ranged else None,
        maxValue=max(numbers) if ranged else None,
        containsDate=True if "date" in kinds else None,
        containsNonDate=False if kinds <= {"date"} else None,
        minDate=min(moments, default=None),
        maxDate=max(moments, default=None),
    )


def cache_item(value: Value) -> CacheItem:
    match value:
        case None:
            return Missing()
        case str():
            return Text(v=value)
        case bool():
            return Boolean(v=value)
        case ExcelError():
            return Error(v=value.code)
        case dt.datetime():
            return DateTimeField(v=value)
        case _:
            return Number(v=value)


def cache_value(item: object) -> Value:
    """The value a shared or inline cache item stands for."""
    match item:
        case Missing():
            return None
        case Error():
            return ExcelError(item.v)
        case Text() | Boolean() | Number() | DateTimeField():
            return item.v
        case _:
            raise InvalidArgumentError(f"Unexpected pivot cache item {item!r}.")


def _records(setup: FieldSetup) -> PivotRecords:
    columns = [
        [SharedIndex(index) for index in setup.shared[position].index]
        if position in setup.shared
        else column.values
        for position, column in enumerate(setup.source.columns)
    ]
    return PivotRecords(list(zip(*columns, strict=True)))
