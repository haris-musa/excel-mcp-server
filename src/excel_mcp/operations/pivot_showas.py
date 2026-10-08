"""The figures a PivotTable shows: summaries, optionally relative to other summaries."""

from excel_mcp.calc.values import DIV0, NA, NULL, ExcelError
from excel_mcp.errors import InvalidArgumentError
from excel_mcp.operations.pivot_axis import Axis, Line, Path
from excel_mcp.operations.pivot_calc import Number
from excel_mcp.operations.pivot_options import ShowAs
from excel_mcp.operations.pivot_values import Aggregator, DataSpec

PERCENT_FORMAT = "0.00%"
PERCENTAGES: frozenset[ShowAs] = frozenset(
    {
        "percent_of_total",
        "percent_of_row",
        "percent_of_column",
        "percent_of_parent_row",
        "percent_of_parent_column",
        "percent_of_parent",
        "percent_difference_from",
        "percent_of",
        "percent_running_total",
    }
)


class _Absent:
    """The place a comparison points at is not a line of the table."""


ABSENT = _Absent()


class Figures:
    """What each cell of the table body shows."""

    def __init__(
        self, aggregator: Aggregator, rows: Axis, columns: Axis, specs: list[DataSpec]
    ) -> None:
        self.aggregator = aggregator
        self.rows, self.columns, self.specs = rows, columns, specs

    def cell(self, row: Line, column: Line) -> Number:
        if row.blank or column.blank:
            return None
        data = row.data if self.rows.value_levels else column.data
        spec = self.specs[data]
        if spec.show_as is None:
            return self._value(row.path, column.path, data)
        return self._shown(spec, row.path, column.path, data)

    def _value(self, rows: Path, columns: Path, data: int) -> Number:
        return self.aggregator.value(rows, columns, data)

    def _shown(self, spec: DataSpec, rows: Path, columns: Path, data: int) -> Number:
        value = self._value(rows, columns, data)
        match spec.show_as:
            case "percent_of_total":
                return _ratio(value, self._value((), (), data))
            case "percent_of_row":
                return _ratio(value, self._value(rows, (), data))
            case "percent_of_column":
                return _ratio(value, self._value((), columns, data))
            case "percent_of_parent_row":
                if not self.rows.fields:
                    return None
                return _share(value, self._value(rows[:-1], columns, data))
            case "percent_of_parent_column":
                if not self.columns.fields:
                    return None
                return _share(value, self._value(rows, columns[:-1], data))
            case _:
                return self._along_field(spec, rows, columns, data, value)

    def _along_field(
        self, spec: DataSpec, rows: Path, columns: Path, data: int, value: Number
    ) -> Number:
        on_rows = any(field.index == spec.base_field for field in self.rows.fields)
        axis = self.rows if on_rows else self.columns
        path = rows if on_rows else columns
        level = next(i for i, field in enumerate(axis.fields) if field.index == spec.base_field)
        if len(path) <= level:
            return None

        def at(item: int) -> Number | _Absent:
            moved = (*path[:level], item, *path[level + 1 :])
            if not axis.has(moved):
                return ABSENT
            return self._value(moved, columns, data) if on_rows else self._value(rows, moved, data)

        if spec.show_as == "percent_of_parent":
            ancestor = path[: level + 1]
            whole = (
                self._value(ancestor, columns, data)
                if on_rows
                else self._value(rows, ancestor, data)
            )
            return _share(value, whole)
        existing = [
            (item, figure)
            for item in axis.children[path[:level]]
            if not isinstance(figure := at(item), _Absent)
        ]
        items, series = [item for item, _ in existing], [figure for _, figure in existing]
        position = items.index(path[level])
        match spec.show_as:
            case "running_total" | "percent_running_total":
                running = _total(series[: position + 1]) or 0.0
                if spec.show_as == "running_total":
                    return running
                parent = path[:level]
                whole = (
                    self._value(parent, columns, data)
                    if on_rows
                    else self._value(rows, parent, data)
                )
                innermost = level == len(axis.fields) - 1
                return _share(running, _whole(spec, innermost, whole, series))
            case "rank_ascending" | "rank_descending":
                return _rank(series, position, spec.show_as == "rank_ascending")
        if spec.base_item is None:
            if position == 0:
                return _first(spec, value)
            return _relative(spec.show_as, value, series[position - 1])
        if path[level] == spec.base_item:
            return None if spec.show_as != "percent_of" or value is None else _ratio(value, value)
        base = at(spec.base_item)
        return NA if isinstance(base, _Absent) else _relative(spec.show_as, value, base)


def _whole(spec: DataSpec, innermost: bool, parent: Number, series: list[Number]) -> Number:
    """What a running total is a share of: the parent's figure, which for an average, minimum
    or calculated field is not the total of the items."""
    if innermost:
        return parent
    if spec.function in ("sum", "count") and spec.key[0] == "column":
        return _total(series)
    raise InvalidArgumentError(
        "show_as 'percent_running_total' of an average, minimum, maximum or calculated field "
        "needs base_field to be the innermost field of its axis."
    )


def _first(spec: DataSpec, value: Number) -> Number:
    """The first item has nothing before it to compare with, except itself."""
    if spec.show_as != "percent_of" or value is None:
        return None
    return _ratio(value, value)


def _num(value: Number) -> float:
    return 0.0 if value is None or isinstance(value, ExcelError) else float(value)


def _ratio(value: Number, whole: Number) -> Number:
    """``value`` as a fraction of ``whole``; Excel reports a missing or zero whole as an error."""
    if isinstance(value, ExcelError):
        return value
    if isinstance(whole, ExcelError):
        return whole
    return _num(value) / whole if whole else DIV0


def _share(value: Number, whole: Number | None) -> Number:
    """Like _ratio, but a missing or zero whole leaves the cell blank."""
    if isinstance(value, ExcelError) or isinstance(whole, ExcelError):
        return value if isinstance(value, ExcelError) else whole
    return _num(value) / whole if whole else None


def _total(values: list[Number]) -> Number:
    """The total of the figures, or None when there are none."""
    figures = [value for value in values if value is not None]
    for value in figures:
        if isinstance(value, ExcelError):
            return value
    return sum(_num(value) for value in figures) if figures else None


def _rank(series: list[Number], position: int, ascending: bool) -> Number:
    """1 for the first, counting ties as one place, as Excel does."""
    current = series[position]
    if current is None or isinstance(current, ExcelError):
        return current
    others = {v for v in series if isinstance(v, int | float)}
    return 1 + sum(1 for value in others if (value < current if ascending else value > current))


def _relative(show_as: ShowAs | None, value: Number, base: Number) -> Number:
    if show_as == "difference_from":
        for operand in (value, base):
            if isinstance(operand, ExcelError):
                return operand
        return _num(value) - _num(base)
    if value is None:
        return NULL
    if base is None:
        return None
    ratio = _ratio(value, base)
    if show_as == "percent_of" or isinstance(ratio, ExcelError) or ratio is None:
        return ratio
    return ratio - 1
