"""Creating PivotTables that Excel can refresh and rearrange."""

from typing import Literal

from openpyxl.pivot.table import (
    DataField,
    FieldItem,
    Location,
    PageField,
    PivotField,
    PivotTableStyle,
    RowColField,
    TableDefinition,
)
from openpyxl.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet
from pydantic import BaseModel, Field

from excel_mcp.errors import InvalidArgumentError
from excel_mcp.operations.cells import stored_cells, writable_cell
from excel_mcp.operations.pivot_cache import FieldItems, build_cache, build_items
from excel_mcp.operations.pivot_index import pivot_area, sheet_pivots, workbook_pivots
from excel_mcp.operations.pivot_layout import (
    CellContent,
    DataSpec,
    Function,
    Shape,
    axis_items,
    axis_slots,
    render,
)
from excel_mcp.operations.pivot_source import Source, read_source
from excel_mcp.refs import CellRange, parse_cell, parse_range

_CAPTIONS = {
    "sum": "Sum",
    "count": "Count",
    "average": "Average",
    "min": "Min",
    "max": "Max",
}
_DATA_FIELD_COLUMNS = -2


class PivotValue(BaseModel):
    field: str = Field(description="Header of the column to summarize.")
    function: Function = Field(
        default="sum", description="'count' counts non-empty cells; the others need numbers."
    )


def create_pivot(
    workbook: Workbook,
    source_sheet: Worksheet,
    source_range: str,
    target_sheet: Worksheet,
    target_cell: str,
    rows: list[str],
    columns: list[str],
    values: list[PivotValue],
    filters: list[str],
    name: str | None,
    max_cells: int,
) -> tuple[str, CellRange]:
    source = read_source(source_sheet, source_range, max_cells)
    row_fields = [source.field_index(field) for field in rows]
    column_fields = [source.field_index(field) for field in columns]
    filter_fields = [source.field_index(field) for field in filters]
    _check_distinct(source, row_fields, column_fields, filter_fields)
    specs = _data_specs(source, values)
    items = {
        field: build_items(source.columns[field])
        for field in [*row_fields, *column_fields, *filter_fields]
    }
    shape = Shape(source, row_fields, column_fields, specs, items)

    row_slots = axis_slots(row_fields, shape)
    column_slots = axis_slots(column_fields, shape, data_count=len(specs))
    filter_height = len(filters) + 1 if filters else 0
    start_row, start_col = parse_cell(target_cell)
    body = CellRange(
        start_row + filter_height,
        start_col,
        start_row + filter_height + shape.header_rows + len(row_slots) - 1,
        start_col + len(row_fields) + len(column_slots) - 1,
    )
    area = CellRange(start_row, start_col, body.max_row, body.max_col).within(max_cells)
    _check_free(target_sheet, area, source_sheet, source.ref)

    pivot_name = _pivot_name(workbook, name)
    cache = build_cache(source, items)
    cache_id = max((pivot.cacheId for pivot in workbook_pivots(workbook)), default=0) + 1
    pivot = TableDefinition(
        name=pivot_name,
        cacheId=cache_id,
        dataCaption="Values",
        updatedVersion=6,
        minRefreshableVersion=3,
        createdVersion=6,
        useAutoFormatting=True,
        itemPrintTitles=True,
        indent=0,
        compact=False,
        compactData=False,
        multipleFieldFilters=False,
        applyWidthHeightFormats=True,
        location=_location(shape, body, len(filters)),
        pivotFields=_pivot_fields(shape, filter_fields),
        rowFields=[RowColField(x=field) for field in row_fields],
        rowItems=axis_items(row_slots, with_data=False),
        colFields=_column_fields(column_fields, len(specs)),
        colItems=axis_items(column_slots, with_data=len(specs) > 1),
        pageFields=[PageField(fld=field, hier=-1) for field in filter_fields],
        dataFields=[_data_field(source, spec) for spec in specs],
        pivotTableStyleInfo=PivotTableStyle(
            name="PivotStyleLight16",
            showRowHeaders=True,
            showColHeaders=True,
            showRowStripes=False,
            showColStripes=False,
            showLastColumn=True,
        ),
    )
    pivot.cache = cache
    target_sheet.add_pivot(pivot)

    cells = render(shape, row_slots, column_slots)
    for (row, column), (value, number_format) in cells.items():
        _write(target_sheet, body.min_row + row, body.min_col + column, value, number_format)
    for position, field in enumerate(filter_fields):
        _write(target_sheet, start_row + position, start_col, source.columns[field].name, None)
        _write(target_sheet, start_row + position, start_col + 1, "(All)", None)
    return pivot_name, body


def _check_distinct(
    source: Source, rows: list[int], columns: list[int], filters: list[int]
) -> None:
    used = [*rows, *columns, *filters]
    repeated = {source.columns[field].name for field in used if used.count(field) > 1}
    if repeated:
        raise InvalidArgumentError(
            f"Each field can be used once in rows, columns and filters; repeated: "
            f"{sorted(repeated)}."
        )


def _data_specs(source: Source, values: list[PivotValue]) -> list[DataSpec]:
    specs = []
    for value in values:
        field = source.field_index(value.field)
        column = source.columns[field]
        if value.function != "count" and column.kind != "number":
            raise InvalidArgumentError(
                f"Field {column.name!r} holds {column.kind}, not numbers, so it cannot be "
                f"{value.function!r}-ed. Use function 'count'."
            )
        caption = f"{_CAPTIONS[value.function]} of {column.name}"
        if any(spec.caption.casefold() == caption.casefold() for spec in specs):
            raise InvalidArgumentError(f"{caption!r} is listed twice in values.")
        if caption.casefold() in {other.name.casefold() for other in source.columns}:
            raise InvalidArgumentError(f"Rename the source column {caption!r}: it clashes.")
        specs.append(DataSpec(field, value.function, caption))
    return specs


def _check_free(
    sheet: Worksheet, area: CellRange, source_sheet: Worksheet, source_ref: str
) -> None:
    if sheet is source_sheet and area.overlaps(parse_range(source_ref)):
        raise InvalidArgumentError(
            f"The PivotTable would cover its own source data {source_ref}. Choose a "
            "target_cell outside it, or another target_sheet."
        )
    for pivot in sheet_pivots(sheet):
        if area.overlaps(pivot_area(pivot)):
            raise InvalidArgumentError(
                f"{area} overlaps the PivotTable {pivot.name!r} at {pivot.location.ref}."
            )
    for cell in stored_cells(sheet):
        if area.min_row <= cell.row <= area.max_row and area.min_col <= cell.column <= area.max_col:
            raise InvalidArgumentError(
                f"The PivotTable would cover {area}, but {cell.coordinate} already holds a "
                "value. Choose an empty target_cell, or clear the range first."
            )


def _pivot_name(workbook: Workbook, name: str | None) -> str:
    existing = {pivot.name.casefold() for pivot in workbook_pivots(workbook)}
    if name is None:
        number = len(existing) + 1
        while f"pivottable{number}" in existing:
            number += 1
        return f"PivotTable{number}"
    if not name.strip():
        raise InvalidArgumentError("name cannot be empty.")
    if name.casefold() in existing:
        raise InvalidArgumentError(f"A PivotTable named {name!r} already exists.")
    return name


def _location(shape: Shape, body: CellRange, filter_count: int) -> Location:
    return Location(
        ref=str(body),
        firstHeaderRow=1,
        firstDataRow=shape.header_rows,
        firstDataCol=len(shape.rows),
        rowPageCount=filter_count or None,
        colPageCount=1 if filter_count else None,
    )


def _pivot_fields(shape: Shape, filters: list[int]) -> list[PivotField]:
    axes: dict[int, Literal["axisRow", "axisCol", "axisPage"]] = {
        **dict.fromkeys(filters, "axisPage"),
        **dict.fromkeys(shape.columns, "axisCol"),
        **dict.fromkeys(shape.rows, "axisRow"),
    }
    summarized = {spec.field for spec in shape.values}
    return [
        PivotField(
            axis=axes.get(position),
            dataField=True if position in summarized else None,
            items=_field_items(shape.items.get(position)),
            numFmtId=column.number_format_id or None,
            compact=False,
            outline=False,
            showAll=False,
        )
        for position, column in enumerate(shape.source.columns)
    ]


def _field_items(items: FieldItems | None) -> list[FieldItem]:
    if items is None:
        return []
    return [*(FieldItem(x=position) for position in items.order), FieldItem(t="default")]


def _column_fields(columns: list[int], value_count: int) -> list[RowColField]:
    fields = [RowColField(x=field) for field in columns]
    if value_count > 1:
        fields.append(RowColField(x=_DATA_FIELD_COLUMNS))
    return fields


def _data_field(source: Source, spec: DataSpec) -> DataField:
    column = source.columns[spec.field]
    return DataField(
        name=spec.caption,
        fld=spec.field,
        subtotal=spec.function,
        baseField=0,
        baseItem=0,
        numFmtId=None if spec.function == "count" else column.number_format_id or None,
    )


def _write(
    sheet: Worksheet,
    row: int,
    column: int,
    value: CellContent,
    number_format: str | None,
) -> None:
    cell = writable_cell(sheet, row, column)
    cell.value = value
    if isinstance(value, str):
        cell.data_type = "s"
    if number_format:
        cell.number_format = number_format
