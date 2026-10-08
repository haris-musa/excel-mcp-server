"""Summarizing the source records that fall under a row group and a column group."""

from dataclasses import dataclass

from excel_mcp.calc.values import ExcelError
from excel_mcp.operations.pivot_axis import Path
from excel_mcp.operations.pivot_calc import CalcField, Evaluator, FieldKey, Number
from excel_mcp.operations.pivot_options import Function, ShowAs
from excel_mcp.operations.pivot_source import Source

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
    """Summaries of record groups, calculated on demand."""

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
            # Excel calculates a calculated field even where no record falls: from sums of 0.
            self.memo[key] = (
                self._summary(spec, records) if records or spec.key[0] == "calculated" else None
            )
        return self.memo[key]

    def _summary(self, spec: DataSpec, records: list[int]) -> Number:
        kind, index = spec.key
        if kind == "calculated":
            return self._calculated(self.calculated[index], records)
        column = self.source.columns[index]
        present = [column.values[record] for record in records if column.values[record] is not None]
        if spec.function == "count":
            return len(present) or None
        numbers = [value for value in present if isinstance(value, int | float)]
        if not numbers:
            return None
        match spec.function:
            case "sum":
                return sum(numbers)
            case "average":
                return sum(numbers) / len(numbers)
            case "min":
                return min(numbers)
            case _:
                return max(numbers)

    def _calculated(self, calc: CalcField, records: list[int]) -> Number:
        def value_of(key: FieldKey) -> Number:
            kind, index = key
            if kind == "calculated":
                return self._calculated(self.calculated[index], records)
            return sum(
                value
                for record in records
                if isinstance(value := self.source.columns[index].values[record], int | float)
            )

        return self.evaluator.evaluate(calc, value_of)


def is_error(value: Number) -> bool:
    return isinstance(value, ExcelError)
