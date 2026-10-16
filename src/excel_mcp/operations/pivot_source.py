"""Reading the data a PivotTable summarizes."""

import datetime as dt
from collections.abc import Callable
from dataclasses import dataclass
from functools import cached_property
from typing import Literal, TypeGuard, cast

from openpyxl.cell.cell import Cell, MergedCell
from openpyxl.utils.datetime import to_excel
from openpyxl.worksheet.worksheet import Worksheet

from excel_mcp.calc.values import ERRORS, ExcelError
from excel_mcp.errors import InvalidArgumentError, LimitExceededError
from excel_mcp.operations.comparison import FormulaResults
from excel_mcp.refs import CellRange, parse_range
from excel_mcp.text import quoted

Value = str | int | float | bool | dt.datetime | ExcelError | None
"""A cell as a PivotTable sees it. Dates, times and durations are moments counted from `EPOCH`."""
Kind = Literal["number", "date", "other"]
Results = Callable[[Worksheet, CellRange], FormulaResults]

EPOCH = dt.datetime(1899, 12, 30)
"""Where the cache counts date serial numbers from; times and durations are dates of 1899-1900."""


def is_number(value: Value) -> TypeGuard[int | float]:
    return isinstance(value, int | float) and not isinstance(value, bool)


def kind_of(value: Value) -> Kind | None:
    if value is None:
        return None
    if is_number(value):
        return "number"
    return "date" if isinstance(value, dt.datetime) else "other"


def serial(moment: dt.datetime) -> float:
    return (moment - EPOCH) / dt.timedelta(days=1)


def moment_of(number: float) -> dt.datetime:
    return EPOCH + dt.timedelta(milliseconds=round(number * 86_400_000))


@dataclass(frozen=True)
class Column:
    name: str
    values: list[Value]
    number_format: str
    number_format_id: int
    """The format of every cell of the column, or General when they differ, as Excel takes it."""

    @cached_property
    def kinds(self) -> set[Kind]:
        return {kind for value in self.values if (kind := kind_of(value))}

    @cached_property
    def has_errors(self) -> bool:
        return any(isinstance(value, ExcelError) for value in self.values)


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
            raise InvalidArgumentError(
                f"Field {name!r} not found. "
                f"Available fields: {quoted(column.name for column in self.columns)}."
            )
        return folded.index(name.strip().casefold())


def read_source(sheet: Worksheet, source_range: str, max_cells: int, results: Results) -> Source:
    """Read the block; ``results`` gives the results of its formulas (when it has any)."""
    area = parse_range(source_range)
    if area.size > max_cells:
        raise LimitExceededError(
            f"Source {area} has {area.size:,} cells; PivotTables can read at most "
            f"{max_cells:,}. Use a smaller source, or raise --max-pivot-cells "
            "(environment variable EXCEL_MCP_MAX_PIVOT_CELLS)."
        )
    if area.rows < 2:
        raise InvalidArgumentError(f"source {area} needs a header row and at least one data row.")
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
    calculated: FormulaResults = {}
    if any(cell.data_type == "f" for row in rows[1:] for cell in row):
        calculated = results(sheet, area)
    columns = [
        _column(name, [row[index] for row in rows[1:]], calculated)
        for index, name in enumerate(names)
    ]
    return Source(sheet.title, str(area), columns)


def _header(cell: Cell | MergedCell) -> str:
    if cell.data_type != "s" or not str(cell.value).strip():
        raise InvalidArgumentError(
            f"Header cell {cell.coordinate} must contain text: the first row of "
            "source is the header row, with one label per column."
        )
    return str(cell.value)


def _column(name: str, cells: list[Cell | MergedCell], calculated: FormulaResults) -> Column:
    formats = {(cell.number_format, cell._style.numFmtId) for cell in cells}
    shared = formats.pop() if len(formats) == 1 else ("General", 0)
    return Column(name, [_value(cell, calculated) for cell in cells], *shared)


def _value(cell: Cell | MergedCell, calculated: FormulaResults) -> Value:
    if cell.data_type == "f":
        return _result(calculated.get(cell.coordinate))
    if cell.data_type == "e":
        return ExcelError(str(cell.value))
    return _plain(cell.value)


def _plain(value: object) -> Value:
    if isinstance(value, dt.datetime | dt.date | dt.time | dt.timedelta):
        return moment_of(to_excel(value))
    return cast(Value, value)


def _result(result: object) -> Value:
    """A calculated result: errors come as their text, dates and times as moments."""
    if isinstance(result, str) and result in ERRORS:
        return ERRORS[result]
    return _plain(result)
