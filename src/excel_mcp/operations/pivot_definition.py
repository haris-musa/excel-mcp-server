"""The PivotTable definition Excel reads: its fields, axes, filters and values."""

from dataclasses import dataclass
from typing import Literal

from openpyxl.pivot.fields import Index
from openpyxl.pivot.table import (
    AutoSortScope,
    DataField,
    FieldItem,
    Location,
    PageField,
    PivotArea,
    PivotField,
    PivotFilter,
    PivotTableStyle,
    Reference,
    RowColField,
    TableDefinition,
)
from openpyxl.utils.datetime import to_excel
from openpyxl.worksheet.filters import AutoFilter, CustomFilter, CustomFilters, FilterColumn

from excel_mcp.operations.pivot_axis import Axis, axis_items
from excel_mcp.operations.pivot_calc import CalcField
from excel_mcp.operations.pivot_fields import AxisField, FieldSetup
from excel_mcp.operations.pivot_options import DatePeriod, Layout, ShowAs
from excel_mcp.operations.pivot_render import Table
from excel_mcp.operations.pivot_values import DataSpec
from excel_mcp.refs import CellRange

VALUES_FIELD = -2
PREVIOUS_ITEM = 1048828
VALUES_REFERENCE = 4294967294
_SHOW_DATA_AS: dict[ShowAs, str] = {
    "percent_of_total": "percentOfTotal",
    "percent_of_row": "percentOfRow",
    "percent_of_column": "percentOfCol",
    "difference_from": "difference",
    "percent_difference_from": "percentDiff",
    "percent_of": "percent",
    "running_total": "runTotal",
}
X14 = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
SHOW_VALUES_AS = "{E15A36E0-9728-4e99-A89B-3F7291B0FE68}"
SHOW_AS_EXTENSION: dict[ShowAs, str] = {
    "percent_of_parent_row": "percentOfParentRow",
    "percent_of_parent_column": "percentOfParentCol",
    "percent_of_parent": "percentOfParent",
    "percent_running_total": "percentOfRunningTotal",
    "rank_ascending": "rankAscending",
    "rank_descending": "rankDescending",
}
_Axis = Literal["axisRow", "axisCol", "axisPage"]


@dataclass(frozen=True)
class PageFilter:
    field: AxisField
    item: int | None
    """The one item shown, when exactly one is."""


@dataclass(frozen=True)
class Definition:
    name: str
    cache_id: int
    setup: FieldSetup
    calculated: list[CalcField]
    rows: Axis
    columns: Axis
    pages: list[PageFilter]
    hidden: list[AxisField]
    """Fields off the axes whose items are limited."""
    filters: list[PivotFilter]
    specs: list[DataSpec]
    format_ids: list[int | None]
    table: Table
    body: CellRange
    layout: Layout
    subtotals: bool


def build_definition(plan: Definition) -> TableDefinition:
    tabular, compact = plan.layout == "tabular", plan.layout == "compact"
    return TableDefinition(
        name=plan.name,
        cacheId=plan.cache_id,
        dataOnRows=True if plan.rows.value_levels else None,
        dataCaption="Values",
        updatedVersion=6,
        minRefreshableVersion=3,
        createdVersion=6,
        useAutoFormatting=True,
        itemPrintTitles=True,
        indent=0,
        compact=compact,
        compactData=compact,
        outline=not tabular,
        outlineData=not tabular,
        multipleFieldFilters=False,
        applyWidthHeightFormats=True,
        location=Location(
            ref=str(plan.body),
            firstHeaderRow=1,
            firstDataRow=plan.table.header_rows,
            firstDataCol=plan.table.label_columns,
            rowPageCount=len(plan.pages) or None,
            colPageCount=1 if plan.pages else None,
        ),
        pivotFields=_pivot_fields(plan),
        rowFields=_axis_fields(plan.rows),
        rowItems=axis_items(plan.rows),
        colFields=_axis_fields(plan.columns),
        colItems=axis_items(plan.columns),
        pageFields=[
            PageField(fld=page.field.index, item=page.item, hier=-1) for page in plan.pages
        ],
        dataFields=[_data_field(plan, position) for position in range(len(plan.specs))],
        filters=plan.filters,
        pivotTableStyleInfo=PivotTableStyle(
            name="PivotStyleLight16",
            showRowHeaders=True,
            showColHeaders=True,
            showRowStripes=False,
            showColStripes=False,
            showLastColumn=True,
        ),
    )


def _axis_fields(axis: Axis) -> list[RowColField]:
    fields = [RowColField(x=field.index) for field in axis.fields]
    if axis.value_levels:
        fields.append(RowColField(x=VALUES_FIELD))
    return fields


def _pivot_fields(plan: Definition) -> list[PivotField]:
    axes: dict[int, _Axis] = {
        **{page.field.index: "axisPage" for page in plan.pages},
        **{field.index: "axisCol" for field in plan.columns.fields},
        **{field.index: "axisRow" for field in plan.rows.fields},
    }
    placed = [*plan.rows.fields, *plan.columns.fields]
    axis_fields = {
        field.index: field for field in [*placed, *(p.field for p in plan.pages), *plan.hidden]
    }
    on_axes = {field.index for field in placed}
    source = plan.setup.source
    summarized = {spec.field for spec in plan.specs}
    fields = []
    for position in range(len(source.columns) + plan.setup.extra_fields):
        column = source.columns[position] if position < len(source.columns) else None
        axis = axis_fields.get(position)
        totals = plan.subtotals or position not in on_axes
        fields.append(
            PivotField(
                axis=axes.get(position),
                dataField=True if position in summarized else None,
                items=_items(plan, position, axis, totals),
                numFmtId=(column.number_format_id or None) if column else None,
                compact=None if plan.layout == "compact" else False,
                outline=False if plan.layout == "tabular" else None,
                showAll=False,
                defaultSubtotal=None if totals else False,
                sortType=_sort_type(axis),
                multipleItemSelectionAllowed=_multiple(plan, axis),
                autoSortScope=_sort_scope(plan, axis),
            )
        )
    for _ in plan.calculated:
        position = len(fields)
        fields.append(
            PivotField(
                dataField=True if position in summarized else None,
                compact=None if plan.layout == "compact" else False,
                outline=False if plan.layout == "tabular" else None,
                dragToRow=False,
                dragToCol=False,
                dragToPage=False,
                showAll=False,
                defaultSubtotal=False,
            )
        )
    return fields


def _items(
    plan: Definition, position: int, axis: AxisField | None, with_default: bool
) -> list[FieldItem]:
    if axis is not None:
        shown = axis.visible
        items = [
            FieldItem(x=item_id, h=True if shown is not None and place not in shown else None)
            for place, item_id in enumerate(axis.item_ids)
        ]
    elif position in plan.setup.shared:
        items = [FieldItem(x=item_id) for item_id in plan.setup.shared[position].order]
    else:
        return []
    return [*items, *([FieldItem(t="default")] if with_default else [])]


def _sort_type(axis: AxisField | None) -> Literal["manual", "ascending", "descending"]:
    if axis is None or not (axis.descending or axis.sort_by) or (axis.ranges and not axis.sort_by):
        return "manual"
    return "descending" if axis.descending else "ascending"


def _multiple(plan: Definition, axis: AxisField | None) -> bool | None:
    page = next((p for p in plan.pages if p.field is axis), None)
    if page is None or not axis or not axis.hides_items or page.item is not None:
        return None
    return True


def _sort_scope(plan: Definition, axis: AxisField | None) -> AutoSortScope | None:
    if axis is None or axis.sort_by is None:
        return None
    data = next(i for i, spec in enumerate(plan.specs) if spec.caption == axis.sort_by)
    return AutoSortScope(
        pivotArea=PivotArea(
            type=None,
            dataOnly=False,
            outline=False,
            fieldPosition=0,
            references=[Reference(field=VALUES_REFERENCE, selected=False, x=[Index(v=data)])],
        )
    )


def _data_field(plan: Definition, position: int) -> DataField:
    spec = plan.specs[position]
    show_as = None
    if spec.show_as in _SHOW_DATA_AS:
        show_as = _SHOW_DATA_AS[spec.show_as]
    base_item = (
        PREVIOUS_ITEM
        if spec.show_as in _PREVIOUS_BY_DEFAULT and spec.base_item is None
        else spec.base_item
    )
    return DataField(
        name=spec.caption,
        fld=spec.field,
        subtotal=spec.function,
        showDataAs=show_as or "normal",
        baseField=spec.base_field or 0,
        baseItem=base_item or 0,
        numFmtId=plan.format_ids[position],
    )


_PREVIOUS_BY_DEFAULT = ("difference_from", "percent_difference_from", "percent_of")


def data_field_extensions(plan: Definition) -> dict[int, str]:
    """The ``<extLst>`` of the data fields whose "show values as" has no attribute of its own
    (rank, percent of parent), by position, for `SheetPackage.pivot_fields`."""
    return {
        position: (
            f'<extLst><ext uri="{SHOW_VALUES_AS}" xmlns:x14="{X14}">'
            f'<x14:dataField pivotShowAs="{SHOW_AS_EXTENSION[spec.show_as]}"/></ext></extLst>'
        )
        for position, spec in enumerate(plan.specs)
        if spec.show_as in SHOW_AS_EXTENSION
    }


def date_filter(field: int, period: DatePeriod, number: int) -> PivotFilter:
    """The "between" date filter Excel writes for a date range."""
    first, last = str(int(to_excel(period.start))), str(int(to_excel(period.end)))
    return PivotFilter(
        fld=field,
        type="dateBetween",
        evalOrder=-1,
        id=number,
        stringValue1=first,
        stringValue2=last,
        autoFilter=AutoFilter(
            ref="A1",
            filterColumn=[
                FilterColumn(
                    colId=0,
                    customFilters=CustomFilters(
                        _and=True,
                        customFilter=[
                            CustomFilter(operator="greaterThanOrEqual", val=first),
                            CustomFilter(operator="lessThanOrEqual", val=last),
                        ],
                    ),
                )
            ],
        ),
    )
