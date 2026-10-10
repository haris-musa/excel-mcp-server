"""Creating PivotTables that Excel can refresh and rearrange."""

from dataclasses import dataclass, field

from openpyxl.styles import Alignment
from openpyxl.styles.numbers import BUILTIN_FORMATS_MAX_SIZE, BUILTIN_FORMATS_REVERSE
from openpyxl.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from excel_mcp.calc.values import ExcelError
from excel_mcp.errors import InvalidArgumentError
from excel_mcp.operations.cells import stored_cells, writable_cell
from excel_mcp.operations.pivot_axis import Axis, Path, Ranker, build_axis
from excel_mcp.operations.pivot_cache import build_cache
from excel_mcp.operations.pivot_calc import CalcField, compile_calculated
from excel_mcp.operations.pivot_definition import (
    Definition,
    PageFilter,
    build_definition,
    data_field_extensions,
    date_filter,
)
from excel_mcp.operations.pivot_fields import AxisField, limit_groups, plan_fields
from excel_mcp.operations.pivot_index import pivot_area, sheet_pivots, workbook_pivots
from excel_mcp.operations.pivot_options import (
    CalculatedField,
    DatePeriod,
    Layout,
    PivotField,
    PivotValue,
    ValuesIn,
)
from excel_mcp.operations.pivot_render import Shown, item_label, number_format, render
from excel_mcp.operations.pivot_showas import Figures
from excel_mcp.operations.pivot_source import Source, read_source
from excel_mcp.operations.pivot_specs import (
    check_values_in_rows,
    data_specs,
    resolve_comparisons,
    sorted_by,
)
from excel_mcp.operations.pivot_values import Aggregator, DataSpec
from excel_mcp.package import state_of
from excel_mcp.refs import CellRange, parse_cell, parse_range


@dataclass(frozen=True)
class PivotRequest:
    rows: list[str]
    columns: list[str]
    values: list[PivotValue]
    filters: list[str]
    fields: list[PivotField]
    calculated: list[CalculatedField]
    layout: Layout
    subtotals: bool
    values_in: ValuesIn
    name: str | None
    filtered: list[PivotField] = field(default_factory=list)
    """Fields outside rows, columns and filters whose items are limited, as slicers do."""
    periods: list[DatePeriod] = field(default_factory=list)
    """Date ranges that records must fall in, as timelines do."""
    shown: dict[str, list[str]] = field(default_factory=dict)
    """The items shown of date groups and of the dates they group, by field name such as
    'Months (Date)'; as slicers choose them."""


def create_pivot(
    workbook: Workbook,
    source_sheet: Worksheet,
    source_range: str,
    target_sheet: Worksheet,
    target_cell: str,
    request: PivotRequest,
    max_cells: int,
    *,
    copy_of_name: bool = False,
) -> tuple[str, CellRange]:
    source = read_source(source_sheet, source_range, max_cells)
    row_fields = [source.field_index(field) for field in request.rows]
    column_fields = [source.field_index(field) for field in request.columns]
    filter_fields = [source.field_index(field) for field in request.filters]
    limited = [source.field_index(f.field) for f in request.filtered]
    _check_distinct(source, row_fields, column_fields, filter_fields, limited)
    calculated = compile_calculated(source, request.calculated)
    options = _options(source, request.fields, [*row_fields, *column_fields, *filter_fields])
    options |= dict(zip(limited, request.filtered, strict=True))
    setup = plan_fields(source, [*row_fields, *column_fields, *filter_fields, *limited], options)
    hidden = [axis for index in limited for axis in setup.axis[index]]
    hidden += limit_groups(setup, request.shown)
    specs = data_specs(source, setup, calculated, request.values)
    on_rows = [axis for index in row_fields for axis in setup.axis[index]]
    on_columns = [axis for index in column_fields for axis in setup.axis[index]]
    pages = [axis for index in filter_fields for axis in setup.axis[index]]
    specs = resolve_comparisons(specs, request.values, on_rows + on_columns)
    if len(specs) > 1 and request.values_in == "rows":
        check_values_in_rows(specs, request.subtotals, len(on_rows))
    on_rows = [sorted_by(field, specs) for field in on_rows]
    on_columns = [sorted_by(field, specs) for field in on_columns]
    rows, columns, aggregator = _axes(
        source, calculated, specs, on_rows, on_columns, [*pages, *hidden], request
    )
    formats = [number_format(spec, _source_format(source, spec)) for spec in specs]
    figures = Figures(aggregator, rows, columns, specs)
    table = render(rows, columns, specs, figures, formats, request.layout)
    start_row, start_col = parse_cell(target_cell)
    filter_height = len(pages) + 1 if pages else 0
    body = CellRange(
        start_row + filter_height,
        start_col,
        start_row + filter_height + table.header_rows + table.rows - 1,
        start_col + table.label_columns + table.columns - 1,
    )
    area = CellRange(start_row, start_col, body.max_row, body.max_col).within(max_cells)
    _check_free(target_sheet, area, source_sheet, source.ref)

    pivot_name = _pivot_name(workbook, request.name, copy_of_name)
    page_filters = [_page_filter(axis) for axis in pages]
    plan = Definition(
        name=pivot_name,
        cache_id=max((pivot.cacheId for pivot in workbook_pivots(workbook)), default=0) + 1,
        setup=setup,
        calculated=calculated,
        rows=rows,
        columns=columns,
        pages=page_filters,
        hidden=hidden,
        filters=[
            date_filter(source.field_index(period.field), period, number)
            for number, period in enumerate(request.periods, 1)
        ],
        specs=specs,
        format_ids=[_format_id(workbook, fmt) for fmt in formats],
        table=table,
        body=body,
        layout=request.layout,
        subtotals=request.subtotals,
    )
    pivot = build_definition(plan)
    pivot.cache = build_cache(setup, calculated)
    target_sheet.add_pivot(pivot)
    for position, extension in data_field_extensions(plan).items():
        state_of(workbook).sheet(target_sheet).pivot_fields[pivot_name, "dataField", position] = (
            extension
        )
    for (row, column), shown in table.cells.items():
        _write(target_sheet, body.min_row + row, body.min_col + column, shown)
    for position, page in enumerate(page_filters):
        _write(target_sheet, start_row + position, start_col, Shown(page.field.name))
        _write(target_sheet, start_row + position, start_col + 1, _page_cell(page))
    return pivot_name, body


def _axes(
    source: Source,
    calculated: list[CalcField],
    specs: list[DataSpec],
    on_rows: list[AxisField],
    on_columns: list[AxisField],
    pages: list[AxisField],
    request: PivotRequest,
) -> tuple[Axis, Axis, Aggregator]:
    """The lines of both axes, from the records the chosen items leave."""
    records = _visible_records(source, [*on_rows, *on_columns, *pages], request.periods)
    row_keys = _keys(source, on_rows)
    column_keys = _keys(source, on_columns)
    aggregator = Aggregator(source, calculated, specs, row_keys, column_keys, records)
    several = len(specs) > 1
    rows = build_axis(
        on_rows,
        {row_keys[record] for record in records},
        _ranker(aggregator, specs, on_rows=True),
        layout=request.layout,
        subtotals=request.subtotals,
        values=len(specs) if several and request.values_in == "rows" else 0,
    )
    columns = build_axis(
        on_columns,
        {column_keys[record] for record in records},
        _ranker(aggregator, specs, on_rows=False),
        layout="tabular",
        subtotals=request.subtotals,
        values=len(specs) if several and request.values_in == "columns" else 0,
    )
    return rows, columns, aggregator


def _source_format(source: Source, spec: DataSpec) -> str:
    kind, index = spec.key
    return source.columns[index].number_format if kind == "column" else "General"


def _options(source: Source, fields: list[PivotField], used: list[int]) -> dict[int, PivotField]:
    options: dict[int, PivotField] = {}
    for option in fields:
        index = source.field_index(option.field)
        if index not in used:
            raise InvalidArgumentError(
                f"Field {option.field!r} in fields is not used in rows, columns or filters."
            )
        if index in options:
            raise InvalidArgumentError(f"Field {option.field!r} is listed twice in fields.")
        options[index] = option
    return options


def _check_distinct(
    source: Source, rows: list[int], columns: list[int], filters: list[int], limited: list[int]
) -> None:
    used = [*rows, *columns, *filters, *limited]
    repeated = {source.columns[field].name for field in used if used.count(field) > 1}
    if repeated:
        raise InvalidArgumentError(
            f"Each field can be used once in rows, columns and filters; repeated: "
            f"{sorted(repeated)}."
        )


def _visible_records(
    source: Source, fields: list[AxisField], periods: list[DatePeriod]
) -> list[int]:
    hiding = [field for field in fields if field.visible is not None]
    ranges = [(source.columns[source.field_index(p.field)].values, p) for p in periods]
    records = [
        record
        for record in range(source.record_count)
        if all(field.positions[record] in (field.visible or ()) for field in hiding)
        and all(p.start <= values[record] <= p.end for values, p in ranges)  # pyright: ignore[reportOperatorIssue]
    ]
    if not records:
        raise InvalidArgumentError("The items chosen leave no data to summarize.")
    return records


def _keys(source: Source, fields: list[AxisField]) -> list[Path]:
    return [
        tuple(field.positions[record] for field in fields) for record in range(source.record_count)
    ]


def _ranker(aggregator: Aggregator, specs: list[DataSpec], *, on_rows: bool) -> Ranker:
    def rank(field: AxisField, prefix: Path, item: int) -> tuple[float, ...]:
        if field.sort_by is None:
            return (item,)
        data = next(i for i, spec in enumerate(specs) if spec.caption == field.sort_by)
        path = (*prefix, item)
        value = aggregator.value(path, (), data) if on_rows else aggregator.value((), path, data)
        number = 0.0 if value is None or isinstance(value, ExcelError) else float(value)
        # Excel orders ties by where the items first appear in the source; descending reverses all.
        appearance = field.item_ids[item]
        return (-number, -appearance) if field.descending else (number, appearance)

    return rank


def _page_filter(axis: AxisField) -> PageFilter:
    shown = axis.visible
    return PageFilter(axis, next(iter(shown)) if shown is not None and len(shown) == 1 else None)


def _page_cell(page: PageFilter) -> Shown:
    if not page.field.hides_items:
        return Shown("(All)")
    if page.item is not None:
        return item_label(page.field, page.item)
    return Shown("(Multiple Items)")


def _format_id(workbook: Workbook, number_format: str | None) -> int | None:
    if number_format is None:
        return None
    if number_format in BUILTIN_FORMATS_REVERSE:
        return BUILTIN_FORMATS_REVERSE[number_format]
    formats = workbook._number_formats  # pyright: ignore[reportAttributeAccessIssue]
    return formats.add(number_format) + BUILTIN_FORMATS_MAX_SIZE


def _check_free(
    sheet: Worksheet, area: CellRange, source_sheet: Worksheet, source_ref: str
) -> None:
    if sheet is source_sheet and area.overlaps(parse_range(source_ref)):
        raise InvalidArgumentError(
            f"The PivotTable would cover its own source data {source_ref}. Choose a "
            "`at` outside it, or another `sheet`."
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
                "value. Choose an empty `at`, or clear the range first."
            )


def _pivot_name(workbook: Workbook, name: str | None, copy_of_name: bool) -> str:
    """``copy_of_name`` allows the name of another PivotTable: Excel's sheet copies share one."""
    existing = {pivot.name.casefold() for pivot in workbook_pivots(workbook)}
    if name is None:
        number = len(existing) + 1
        while f"pivottable{number}" in existing:
            number += 1
        return f"PivotTable{number}"
    if not name.strip():
        raise InvalidArgumentError("name cannot be empty.")
    if name.casefold() in existing and not copy_of_name:
        raise InvalidArgumentError(f"A PivotTable named {name!r} already exists.")
    return name


def _write(sheet: Worksheet, row: int, column: int, shown: Shown) -> None:
    cell = writable_cell(sheet, row, column)
    value = shown.value
    if isinstance(value, ExcelError):
        cell.value = value.code
        cell.data_type = "e"
    else:
        cell.value = value
        if isinstance(value, str):
            cell.data_type = "s"
    if shown.number_format:
        cell.number_format = shown.number_format
    if shown.indent:
        cell.alignment = Alignment(horizontal="left", indent=shown.indent)
