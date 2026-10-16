"""Summarizing the source records that fall under a row group and a column group."""

import datetime as dt
from collections.abc import Iterable
from dataclasses import dataclass

from excel_mcp.calc.values import DIV0, ExcelError
from excel_mcp.operations.pivot_axis import Path
from excel_mcp.operations.pivot_calc import CalcField, Evaluator, FieldKey, Number
from excel_mcp.operations.pivot_options import Function, ShowAs
from excel_mcp.operations.pivot_source import Column, Source, Value, is_number, serial

GroupKey = tuple[Path, Path]


@dataclass(frozen=True)
class DataSpec:
    """One field summarized in the table's body."""

    key: FieldKey
    field: int
    """Position among the PivotTable's fields."""
    function: Function
    caption: str
    number_format: str | None
    show_as: ShowAs | None = None
    base_field: int | None = None
    """Position of the field show_as compares along, if it needs one."""
    base_item: int | None = None
    """Position of the item compared with; None means the previous one."""


class Aggregator:
    """Summaries of record groups, calculated on demand.

    Like Excel, sum, average, min and max use the numbers and dates of a group (as serial
    numbers) and skip text, booleans and blanks; count counts all that are not blank. An error
    among the values is the result instead, but not in the count of a single line of figures:
    a total of lines (a subtotal or Grand Total) counts as that error too.
    """

    def __init__(
        self,
        source: Source,
        calculated: list[CalcField],
        specs: list[DataSpec],
        row_keys: list[Path],
        column_keys: list[Path],
        records: list[int],
    ) -> None:
        self.source = source
        self.calculated = calculated
        self.specs = specs
        self.row_keys = row_keys
        self.column_keys = column_keys
        self.depth = (len(row_keys[0]), len(column_keys[0]))
        self.evaluator = Evaluator()
        self.groups: dict[GroupKey, list[int]] = {}
        self.memo: dict[tuple[GroupKey, int], Number] = {}
        for record in records:
            rows, columns = row_keys[record], column_keys[record]
            for row_depth in range(len(rows) + 1):
                for column_depth in range(len(columns) + 1):
                    key = (rows[:row_depth], columns[:column_depth])
                    self.groups.setdefault(key, []).append(record)

    def value(self, rows: Path, columns: Path, data: int) -> Number:
        key = ((rows, columns), data)
        if key not in self.memo:
            records = self.groups.get((rows, columns), [])
            spec = self.specs[data]
            total = (len(rows), len(columns)) != self.depth
            # Excel calculates a calculated field even where no record falls: from sums of 0.
            self.memo[key] = (
                self._summary(spec, records, total)
                if records or spec.key[0] == "calculated"
                else None
            )
        return self.memo[key]

    def _summary(self, spec: DataSpec, records: list[int], total: bool) -> Number:
        kind, index = spec.key
        if kind == "calculated":
            return self._calculated(self.calculated[index], records, total)
        column = self.source.columns[index]
        error = self._error(column, records, total)
        present = [column.values[record] for record in records if column.values[record] is not None]
        if spec.function == "count":
            if error and total:
                return error
            return len(present) or None
        if not present:
            return None
        if error:
            return error
        numbers = _numbers(present)
        match spec.function:
            case "sum":
                return sum(numbers)
            case "average":
                return sum(numbers) / len(numbers) if numbers else DIV0
            case "min":
                return min(numbers, default=0)
            case _:
                return max(numbers, default=0)

    def _error(
        self, column: Column, records: list[int], total: bool, *, last: bool = False
    ) -> ExcelError | None:
        """The first error among the records (the last for a calculated field), in the order
        of the source for a line and in the order the lines are shown for a total."""
        if not column.has_errors:
            return None
        if total:
            records = sorted(records, key=lambda r: (self.row_keys[r], self.column_keys[r]))
        errors = [value for r in records if isinstance(value := column.values[r], ExcelError)]
        if not errors:
            return None
        return errors[-1] if last else errors[0]

    def _calculated(self, calc: CalcField, records: list[int], total: bool) -> Number:
        def value_of(key: FieldKey) -> Number:
            kind, index = key
            if kind == "calculated":
                return self._calculated(self.calculated[index], records, total)
            column = self.source.columns[index]
            error = self._error(column, records, total, last=True)
            if error:
                return error
            return sum(_numbers(column.values[record] for record in records))

        return self.evaluator.evaluate(calc, value_of)


def _numbers(values: Iterable[Value]) -> list[float]:
    """The numbers among the values, with dates as serial numbers."""
    numbers: list[float] = []
    for value in values:
        if isinstance(value, dt.datetime):
            numbers.append(serial(value))
        elif is_number(value):
            numbers.append(value)
    return numbers


def is_error(value: Number) -> bool:
    return isinstance(value, ExcelError)
