"""The cells Excel shows for a PivotTable: headers, item labels and figures."""

import datetime as dt
from dataclasses import dataclass

from excel_mcp.calc.values import ExcelError
from excel_mcp.operations.pivot_axis import Axis, Line, axis_items
from excel_mcp.operations.pivot_fields import AxisField
from excel_mcp.operations.pivot_options import Layout
from excel_mcp.operations.pivot_showas import PERCENT_FORMAT, PERCENTAGES, Figures
from excel_mcp.operations.pivot_values import DataSpec

CellValue = str | int | float | dt.datetime | ExcelError | None


@dataclass(frozen=True)
class Shown:
    value: CellValue
    number_format: str | None = None
    indent: int = 0


@dataclass(frozen=True)
class Table:
    """The table body's size and cells, positioned from its top left corner."""

    cells: dict[tuple[int, int], Shown]
    header_rows: int
    label_columns: int
    rows: int
    columns: int


def render(
    rows: Axis,
    columns: Axis,
    specs: list[DataSpec],
    figures: Figures,
    formats: list[str | None],
    layout: Layout,
) -> Table:
    """``formats`` holds each values field's number format."""
    label_columns = 1 if layout == "compact" else len(rows.fields) + rows.value_levels
    header_rows = 1 + len(columns.fields) + columns.value_levels
    table = Table({}, header_rows, label_columns, len(rows.lines), len(columns.lines))
    cells = table.cells
    _headers(cells, rows, columns, specs, layout, header_rows, label_columns)
    _column_labels(cells, columns, specs, label_columns)
    row_items = axis_items(rows)
    for offset, (line, item) in enumerate(zip(rows.lines, row_items, strict=True)):
        top = header_rows + offset
        _row_labels(cells, rows, line, item.r, top, specs, layout)
        for position, column in enumerate(columns.lines):
            value = figures.cell(line, column)
            if value is not None:
                data = line.data if rows.value_levels else column.data
                cells[(top, label_columns + position)] = Shown(value, formats[data])
    return table


def _headers(
    cells: dict[tuple[int, int], Shown],
    rows: Axis,
    columns: Axis,
    specs: list[DataSpec],
    layout: Layout,
    header_rows: int,
    label_columns: int,
) -> None:
    last = header_rows - 1
    if layout == "compact":
        cells[(last, 0)] = Shown("Row Labels")
    else:
        for level, field in enumerate(rows.fields):
            cells[(last, level)] = Shown(field.name)
        if rows.value_levels:
            cells[(last, len(rows.fields))] = Shown("Values")
    if not columns.fields and not columns.value_levels:
        if len(specs) == 1:
            cells[(0, label_columns)] = Shown(specs[0].caption)
        return
    if not columns.value_levels and len(specs) == 1:
        cells[(0, 0)] = Shown(specs[0].caption)
    if layout == "compact":
        cells[(0, label_columns)] = Shown("Column Labels" if columns.fields else "Values")
        return
    for position, field in enumerate(columns.fields):
        cells[(0, label_columns + position)] = Shown(field.name)
    if columns.value_levels:
        cells[(0, label_columns + len(columns.fields))] = Shown("Values")


def _column_labels(
    cells: dict[tuple[int, int], Shown], columns: Axis, specs: list[DataSpec], first: int
) -> None:
    items = axis_items(columns)
    for offset, (line, item) in enumerate(zip(columns.lines, items, strict=True)):
        position = first + offset
        caption = specs[line.data].caption
        suffix = caption if columns.value_levels else "Total"
        match line.kind:
            case "grand":
                text = f"Total {caption}" if columns.value_levels else "Grand Total"
                cells[(1, position)] = Shown(text)
            case "subtotal":
                level = len(line.path)
                label = item_label(columns.fields[level - 1], line.path[level - 1])
                cells[(level, position)] = Shown(f"{_text(label.value)} {suffix}")
            case _:
                for level in range(item.r, len(columns.fields)):
                    cells[(1 + level, position)] = item_label(
                        columns.fields[level], line.path[level]
                    )
                if columns.value_levels:
                    cells[(1 + len(columns.fields), position)] = Shown(caption)


def _row_labels(
    cells: dict[tuple[int, int], Shown],
    rows: Axis,
    line: Line,
    repeated: int,
    top: int,
    specs: list[DataSpec],
    layout: Layout,
) -> None:
    fields = rows.fields
    compact = layout == "compact"
    caption = specs[line.data].caption
    match line.kind:
        case "grand":
            text = f"Total {caption}" if rows.value_levels else "Grand Total"
            cells[(top, 0)] = Shown(text)
        case "subtotal":
            level = len(line.path) - 1
            label = item_label(fields[level], line.path[level])
            suffix = caption if rows.value_levels else "Total"
            cells[(top, 0 if compact else level)] = Shown(f"{_text(label.value)} {suffix}")
        case "group":
            level = len(line.path) - 1
            cells[(top, 0 if compact else level)] = _indented(
                item_label(fields[level], line.path[level]), level, compact
            )
        case _ if layout == "tabular":
            for level in range(repeated, len(fields)):
                cells[(top, level)] = item_label(fields[level], line.path[level])
            if rows.value_levels:
                cells[(top, len(fields))] = Shown(caption)
        case _:
            last = len(fields) - 1
            if not rows.value_levels:
                cells[(top, 0 if compact else last)] = _indented(
                    item_label(fields[last], line.path[last]), last, compact
                )
            else:
                cells[(top, 0 if compact else len(fields))] = Shown(
                    caption, indent=len(fields) if compact else 0
                )


def _indented(label: Shown, level: int, compact: bool) -> Shown:
    return Shown(label.value, label.number_format, level if compact else 0)


def item_label(field: AxisField, position: int) -> Shown:
    value = field.labels[position]
    if value is None:
        return Shown("(blank)")
    return Shown(value, field.number_format)


def _text(value: CellValue) -> str:
    match value:
        case dt.datetime() if value.time() == dt.time():
            return value.date().isoformat()
        case dt.datetime():
            return value.isoformat(sep=" ")
        case float() if value.is_integer():
            return str(int(value))
        case _:
            return str(value)


def number_format(spec: DataSpec, source_format: str) -> str | None:
    """The format of a values field's figures: the one asked for, or Excel's default."""
    if spec.number_format is not None:
        return spec.number_format
    if spec.show_as in PERCENTAGES:
        return PERCENT_FORMAT
    if spec.show_as is not None:
        return None if spec.show_as.startswith("rank") else _plain(source_format)
    return None if spec.function == "count" else _plain(source_format)


def _plain(source_format: str) -> str | None:
    return None if source_format == "General" else source_format
