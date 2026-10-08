"""The lines of a PivotTable's row or column axis, in the order Excel shows them."""

from collections.abc import Callable
from dataclasses import dataclass
from typing import Literal

from openpyxl.pivot.fields import Index
from openpyxl.pivot.table import RowColItem

from excel_mcp.operations.pivot_fields import AxisField
from excel_mcp.operations.pivot_options import Layout

Path = tuple[int, ...]
LineKind = Literal["leaf", "group", "subtotal", "grand"]
_ITEM_TYPE: dict[LineKind, Literal["data", "default", "grand"]] = {
    "leaf": "data",
    "group": "data",
    "subtotal": "default",
    "grand": "grand",
}


@dataclass(frozen=True)
class Line:
    """One row or column of the table body.

    ``path`` holds the item position at each field level down to this line. A "group" line
    stands for a parent whose children follow, and a "subtotal" for one whose children
    came before. ``data`` is the values field of the line, if the axis lists the values.
    """

    kind: LineKind
    path: Path
    data: int = 0
    blank: bool = False
    """Whether the line carries no figures of its own."""


@dataclass(frozen=True)
class Axis:
    fields: list[AxisField]
    lines: list[Line]
    values: int
    """How many values fields the axis lists, or 0 when it does not list them."""
    children: dict[Path, list[int]]
    """The items under each parent, in display order."""

    @property
    def value_levels(self) -> bool:
        return self.values > 0

    def has(self, path: Path) -> bool:
        """Whether some line of the axis goes through ``path``."""
        return all(path[depth] in self.children.get(path[:depth], ()) for depth in range(len(path)))


Ranker = Callable[[AxisField, Path, int], tuple[float, ...]]
"""Sort key of the item at ``path`` extended by the given position."""


def build_axis(
    fields: list[AxisField],
    leaves: set[Path],
    ranker: Ranker,
    *,
    layout: Layout,
    subtotals: bool,
    values: int,
) -> Axis:
    children = _order(fields, leaves, ranker)
    lines = _lines(fields, children, layout, subtotals, values)
    return Axis(fields, lines, values, children)


def _order(fields: list[AxisField], leaves: set[Path], ranker: Ranker) -> dict[Path, list[int]]:
    found: dict[Path, set[int]] = {}
    for leaf in leaves:
        for depth in range(len(leaf)):
            found.setdefault(leaf[:depth], set()).add(leaf[depth])
    ordered = {}
    for prefix, items in found.items():
        field = fields[len(prefix)]
        ordered[prefix] = sorted(items, key=lambda item: ranker(field, prefix, item))
    return ordered


def _lines(
    fields: list[AxisField],
    children: dict[Path, list[int]],
    layout: Layout,
    subtotals: bool,
    values: int,
) -> list[Line]:
    if not fields:
        return [Line("leaf", (), data) for data in range(max(values, 1))]
    lines: list[Line] = []
    nested = layout != "tabular"

    def emit(prefix: Path) -> None:
        for item in children[prefix]:
            path = (*prefix, item)
            innermost = len(path) == len(fields)
            if innermost:
                if nested and values:
                    lines.append(Line("group", path, blank=True))
                lines.extend(Line("leaf", path, data) for data in range(max(values, 1)))
                continue
            if nested:
                lines.append(Line("group", path, blank=values > 0 or not subtotals))
            emit(path)
            if subtotals and (not nested or values):
                lines.extend(Line("subtotal", path, data) for data in range(max(values, 1)))

    emit(())
    lines.extend(Line("grand", (), data) for data in range(max(values, 1)))
    return lines


def axis_items(axis: Axis) -> list[RowColItem]:
    """Describe the lines the way a PivotTable definition lists them."""
    items = []
    previous: Path = ()
    for line in axis.lines:
        path = line.path + ((line.data,) if axis.value_levels and line.kind == "leaf" else ())
        shared = 0
        while shared < min(len(previous), len(path)) and previous[shared] == path[shared]:
            shared += 1
        repeated = min(shared, max(len(path) - 1, 0))
        entries = [Index(v=position) for position in path[repeated:]]
        if line.kind == "grand":
            entries = [Index(v=0)]
        items.append(RowColItem(t=_ITEM_TYPE[line.kind], r=repeated, i=line.data, x=entries))
        previous = path
    return items
