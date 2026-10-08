"""Reading the data a PivotTable summarizes."""

import datetime as dt
from dataclasses import dataclass
from typing import Literal

from openpyxl.cell.cell import Cell, MergedCell
from openpyxl.worksheet.worksheet import Worksheet

from excel_mcp.errors import InvalidArgumentError
from excel_mcp.refs import parse_range

Value = str | int | float | dt.datetime | None
Kind = Literal["text", "number", "date"]


@dataclass(frozen=True)
class Column:
    name: str
    kind: Kind
    values: list[Value]
    number_format: str
    number_format_id: int


@dataclass(frozen=True)
class Source:
    sheet: str
    ref: str
    columns: list[Column]

    @property
    def record_count(self) -> int:
        return len(self.columns[0].values)

    def field_index(self, name: str) -> int:
        folded = [column.name.strip().casefold() for column in self.columns]
        if name.strip().casefold() not in folded:
            available = [column.name for column in self.columns]
            raise InvalidArgumentError(f"Field {name!r} not found. Available fields: {available}.")
        return folded.index(name.strip().casefold())


def read_source(sheet: Worksheet, source_range: str, max_cells: int) -> Source:
    area = parse_range(source_range).within(max_cells)
    if area.rows < 2:
        raise InvalidArgumentError(
            f"source_range {area} needs a header row and at least one data row."
        )
    rows = [
        list(row)
        for row in sheet.iter_rows(
            min_row=area.min_row, max_row=area.max_row, min_col=area.min_col, max_col=area.max_col
        )
    ]
    names = [_header(cell) for cell in rows[0]]
    folded = [name.strip().casefold() for name in names]
    if len(set(folded)) < len(folded):
        raise InvalidArgumentError(f"Header names must be unique, got {names}.")
    columns = [_column(name, [row[index] for row in rows[1:]]) for index, name in enumerate(names)]
    return Source(sheet.title, str(area), columns)


def _header(cell: Cell | MergedCell) -> str:
    if cell.data_type != "s" or not str(cell.value).strip():
        raise InvalidArgumentError(
            f"Header cell {cell.coordinate} must contain text: the first row of "
            "source_range is the header row, with one label per column."
        )
    return str(cell.value)


def _column(name: str, cells: list[Cell | MergedCell]) -> Column:
    values = [_value(cell) for cell in cells]
    kinds: set[Kind] = {_kind(value) for value in values if value is not None}
    if len(kinds) > 1:
        raise InvalidArgumentError(
            f"Column {name!r} mixes {' and '.join(sorted(kinds))} values. Clean it, or "
            "leave it out of source_range."
        )
    first = next((cell for cell in cells if cell.value is not None), None)
    return Column(
        name=name,
        kind=kinds.pop() if kinds else "text",
        values=values,
        number_format=first.number_format if first else "General",
        number_format_id=first._style.numFmtId if first else 0,
    )


def _value(cell: Cell | MergedCell) -> Value:
    match cell.value:
        case None:
            return None
        case _ if cell.data_type in "feb":
            raise InvalidArgumentError(
                f"Cell {cell.coordinate} holds a formula, error or boolean. PivotTables here "
                "need plain text, numbers and dates: write the values, not formulas."
            )
        case dt.datetime() as moment:
            return moment
        case dt.date() as day:
            return dt.datetime.combine(day, dt.time())
        case str() | int() | float() as plain:
            return plain
        case _:
            raise InvalidArgumentError(f"Cell {cell.coordinate} holds a time or duration.")


def _kind(value: Value) -> Kind:
    if isinstance(value, str):
        return "text"
    return "date" if isinstance(value, dt.datetime) else "number"
