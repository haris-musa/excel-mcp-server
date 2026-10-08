"""The fields of a PivotTable: their items, and which item each source record belongs to."""

import datetime as dt
from dataclasses import dataclass, field, replace

from excel_mcp.errors import InvalidArgumentError
from excel_mcp.operations.pivot_groups import Grouping, group_dates, group_numbers
from excel_mcp.operations.pivot_options import PivotField
from excel_mcp.operations.pivot_source import Column, Source, Value


@dataclass(frozen=True)
class Shared:
    """A column's distinct values in order of first appearance, as the cache stores them."""

    values: list[Value]
    index: list[int]
    """For each record, its value's position in ``values``."""
    order: list[int]
    """Positions in ``values`` from the first item to show to the last."""
    rank: list[int]
    """For each position in ``values``, where it comes in ``order``."""


@dataclass(frozen=True)
class AxisField:
    """A field that can be used in rows, columns or filters, with its items in display order."""

    index: int
    """Position among the PivotTable's fields."""
    name: str
    labels: list[Value]
    positions: list[int]
    """For each record, the position of its item in ``labels``."""
    item_ids: list[int]
    """What each item's entry in the field's item list refers to."""
    number_format: str | None
    descending: bool = False
    sort_by: str | None = None
    visible: frozenset[int] | None = None
    """Positions of the items shown, when some are hidden."""
    ranges: bool = False
    """Whether the items are number ranges or days, which Excel sorts as text."""
    ends: bool = False
    """Whether the first and last item are those for everything outside the groups."""
    levels: int = 1
    """How many date groups the source field was split into, when this is one of them."""

    @property
    def hides_items(self) -> bool:
        return self.visible is not None and len(self.visible) < len(self.labels)


@dataclass
class FieldSetup:
    """Everything the cache and the table need to know about the source's fields."""

    source: Source
    shared: dict[int, Shared] = field(default_factory=dict)
    number_groups: dict[int, Grouping] = field(default_factory=dict)
    date_groups: dict[int, list[tuple[int, Grouping]]] = field(default_factory=dict)
    """Source field to its groups, smallest unit first, each with its cache field index."""
    axis: dict[int, list[AxisField]] = field(default_factory=dict)
    """Source field to the fields it contributes to an axis, outermost first."""
    extra_fields: int = 0
    """How many fields beyond the source's columns the cache has."""


def distinct(column_values: list[Value]) -> Shared:
    seen: dict[object, int] = {}
    values: list[Value] = []
    index: list[int] = []
    for value in column_values:
        key = value.casefold() if isinstance(value, str) else value
        if key not in seen:
            seen[key] = len(values)
            values.append(value)
        index.append(seen[key])
    order = sorted(range(len(values)), key=lambda item: sort_key(values[item]))
    rank = [0] * len(order)
    for position, item in enumerate(order):
        rank[item] = position
    return Shared(values, index, order, rank)


def sort_key(value: Value) -> tuple[bool, object]:
    if value is None:
        return True, 0
    return False, value.casefold() if isinstance(value, str) else value


def plan_fields(source: Source, used: list[int], options: dict[int, PivotField]) -> FieldSetup:
    """Work out the axis fields for the source fields in ``used``."""
    setup = FieldSetup(source)
    for index in used:
        column, option = source.columns[index], options.get(index, PivotField(field=""))
        if option.group_dates and option.group_numbers:
            raise InvalidArgumentError(f"Field {column.name!r} cannot be grouped twice.")
        if option.group_dates:
            _date_fields(setup, index, column, option)
        elif option.group_numbers:
            _number_field(setup, index, column, option)
        else:
            _plain_field(setup, index, column, option)
    return setup


def _plain_field(setup: FieldSetup, index: int, column: Column, option: PivotField) -> None:
    shared = distinct(column.values)
    setup.shared[index] = shared
    axis = AxisField(
        index,
        column.name,
        [shared.values[item] for item in shared.order],
        [shared.rank[item] for item in shared.index],
        shared.order,
        None if column.number_format == "General" else column.number_format,
    )
    setup.axis[index] = [_with_options(axis, option)]


def _number_field(setup: FieldSetup, index: int, column: Column, option: PivotField) -> None:
    if column.kind != "number" or None in column.values:
        raise InvalidArgumentError(
            f"Field {column.name!r} cannot be grouped into ranges: it must hold only numbers, "
            "with no blanks."
        )
    assert option.group_numbers is not None
    grouping = group_numbers([float(value) for value in column.values], option.group_numbers)  # pyright: ignore[reportArgumentType]
    setup.shared[index] = distinct(column.values)
    setup.number_groups[index] = grouping
    axis = AxisField(
        index,
        column.name,
        list(grouping.labels),
        grouping.positions,
        list(range(len(grouping.labels))),
        None,
        ranges=True,
        ends=True,
    )
    setup.axis[index] = [_with_options(axis, option)]


def _date_fields(setup: FieldSetup, index: int, column: Column, option: PivotField) -> None:
    if column.kind != "date" or None in column.values:
        raise InvalidArgumentError(
            f"Field {column.name!r} cannot be grouped by date: it must hold only dates, "
            "with no blanks."
        )
    if option.show_items:
        raise InvalidArgumentError(
            "show_items cannot be combined with date groups; filter the groups."
        )
    values = [value for value in column.values if isinstance(value, dt.datetime)]
    groupings = group_dates(values, option.group_dates)
    setup.shared[index] = distinct(column.values)
    levels = []
    for grouping in reversed(groupings):
        position = len(setup.source.columns) + setup.extra_fields
        setup.extra_fields += 1
        setup.date_groups.setdefault(index, []).append((position, grouping))
        axis = AxisField(
            position,
            f"{grouping.by.capitalize()} ({column.name})",
            list(grouping.labels),
            grouping.positions,
            list(range(len(grouping.labels))),
            None,
            ranges=grouping.by == "days",
            ends=True,
            levels=len(groupings),
        )
        levels.append(_with_options(axis, option))
    setup.axis[index] = list(reversed(levels))


def _with_options(axis: AxisField, option: PivotField) -> AxisField:
    visible = None
    if option.show_items:
        texts = {
            item_text(label).casefold(): position for position, label in enumerate(axis.labels)
        }
        missing = [item for item in option.show_items if item.strip().casefold() not in texts]
        if missing:
            shown = [item_text(label) for label in axis.labels][:20]
            raise InvalidArgumentError(
                f"Field {axis.name!r} has no item {missing[0]!r}. Items: {shown}."
            )
        visible = frozenset(texts[item.strip().casefold()] for item in option.show_items)
    field = replace(
        axis, descending=option.sort == "descending", sort_by=option.sort_by, visible=visible
    )
    return _reversed(field) if field.descending and field.sort_by is None else field


def _reversed(axis: AxisField) -> AxisField:
    """Excel lists the items of a field sorted by label in the order they are shown.

    The items for everything outside a group's range stay last.
    """
    last = len(axis.labels) - 1
    inner = list(range(1, last)) if axis.ends else list(range(last + 1))
    order = [*inner[::-1], *([last, 0] if axis.ends else [])]
    place = {old: new for new, old in enumerate(order)}
    return replace(
        axis,
        labels=[axis.labels[old] for old in order],
        positions=[place[position] for position in axis.positions],
        item_ids=[axis.item_ids[old] for old in order],
        visible=None if axis.visible is None else frozenset(place[p] for p in axis.visible),
    )


def item_text(label: Value) -> str:
    match label:
        case None:
            return "(blank)"
        case dt.datetime() if label.time() == dt.time():
            return label.date().isoformat()
        case dt.datetime():
            return label.isoformat(sep=" ")
        case float() if label.is_integer():
            return str(int(label))
        case _:
            return str(label)
