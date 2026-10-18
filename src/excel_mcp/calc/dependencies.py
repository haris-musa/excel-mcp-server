"""Which cells a formula uses and which formulas use a cell, as the calculator reads formulas."""

from collections import defaultdict, deque
from collections.abc import Iterable, Iterator
from dataclasses import dataclass
from typing import Any

from openpyxl.worksheet.formula import ArrayFormula
from openpyxl.worksheet.worksheet import Worksheet

from excel_mcp.calc.cells import CellKey, stored_in
from excel_mcp.calc.engine import Engine, Position
from excel_mcp.calc.parser import Call, Node, parse
from excel_mcp.calc.tablerefs import bind
from excel_mcp.calc.values import FormulaError, UncalculableError
from excel_mcp.errors import LimitExceededError
from excel_mcp.lazy_workbook import extent
from excel_mcp.refs import MAX_COLUMN, MAX_ROW, parse_range

Rect = tuple[str, int, int, int, int]  # sheet title, top, left, bottom, right
MAX_CHECKS = 5_000_000


@dataclass
class Uses:
    rects: list[Rect]
    unresolved: bool  # an INDIRECT, OFFSET or spill reference that cannot be followed


def formula_text(cell: Any) -> str:
    return str(cell.value.text if isinstance(cell.value, ArrayFormula) else cell.value)


def formula_tree(engine: Engine, cell: Any) -> Node:
    """The syntax tree of a formula cell, with table references replaced by their cells."""
    if isinstance(cell.value, ArrayFormula):
        return bind(
            parse(str(cell.value.text)), engine.workbook, cell.parent, cell.row, cell.column
        )
    return engine.tree(cell)


def uses(engine: Engine, sheet: Worksheet, row: int, col: int) -> Uses:
    """What the formula at this position refers to. Raises when the calculator cannot read it."""
    outer, engine.here = engine.here, Position(sheet, row, col)
    try:
        tree = formula_tree(engine, sheet._cells[(row, col)])
        unresolved: list[Call] = []
        rects = [
            (target.title, top or 1, left or 1, bottom or MAX_ROW, right or MAX_COLUMN)
            for target, top, left, bottom, right in engine.scheduler.rectangles(
                tree, sheet, 0, unresolved
            )
        ]
        return Uses(rects, bool(unresolved))
    finally:
        engine.here = outer


class Dependents:
    """Every formula cell of a workbook with the cells it refers to, looked up the other way."""

    def __init__(self, engine: Engine, max_formulas: int) -> None:
        self.points: dict[tuple[str, int, int], list[CellKey]] = defaultdict(list)
        self.areas: list[tuple[Rect, CellKey]] = []
        self.unknown: list[CellKey] = []
        self.arrays: dict[CellKey, Rect] = {}  # array formulas that fill more than their cell
        self.checks = 0
        count = 0
        for sheet in engine.sheets.values():
            for (row, col), cell in stored_in(sheet, 1, 1, *extent(sheet)):
                if cell.data_type != "f":
                    continue
                count += 1
                if count > max_formulas:
                    raise LimitExceededError(
                        f"The workbook has more than {max_formulas:,} formulas, too many to "
                        "search. Raise --max-cells, or trace precedents instead."
                    )
                self._add(engine, sheet, row, col)

    def _add(self, engine: Engine, sheet: Worksheet, row: int, col: int) -> None:
        key = (sheet.title, row, col)
        formula = sheet._cells[(row, col)].value
        if isinstance(formula, ArrayFormula) and (area := parse_range(formula.ref)).size > 1:
            self.arrays[key] = (key[0], area.min_row, area.min_col, area.max_row, area.max_col)
        try:
            found = uses(engine, sheet, row, col)
        except (UncalculableError, FormulaError):
            self.unknown.append(key)
            return
        if found.unresolved:
            self.unknown.append(key)
        for rect in found.rects:
            if rect[1] == rect[3] and rect[2] == rect[4]:
                self.points[(rect[0], rect[1], rect[2])].append(key)
            else:
                self.areas.append((rect, key))

    def of(self, rects: Iterable[Rect], *, all_levels: bool, limit: int) -> list[CellKey]:
        """The formula cells that use the cells in ``rects``, and with ``all_levels``, those
        that use them in turn. More than ``limit`` of them is an error."""
        found: dict[CellKey, None] = {}
        frontier = list(rects)
        while frontier:
            following: list[Rect] = []
            for rect in frontier:
                for key in self._users(rect):
                    if key not in found:
                        found[key] = None
                        following.append(self.span(key))
                if len(found) > limit:
                    raise LimitExceededError(
                        f"More than {limit:,} formulas depend on this. Ask for direct "
                        "dependents of a smaller range."
                    )
            frontier = following if all_levels else []
        return list(found)

    def span(self, key: CellKey) -> Rect:
        """The cells a formula fills: its own, or the range an array formula spills into."""
        return self.arrays.get(key, (key[0], key[1], key[2], key[1], key[2]))

    def _users(self, rect: Rect) -> Iterator[CellKey]:
        sheet, top, left, bottom, right = rect
        self.checks += len(self.areas)
        if self.checks > MAX_CHECKS:
            raise LimitExceededError("Too much work to find the dependents. Use a smaller range.")
        if (bottom - top + 1) * (right - left + 1) <= len(self.points):
            for row in range(top, bottom + 1):
                for col in range(left, right + 1):
                    yield from self.points.get((sheet, row, col), ())
        else:
            for (name, row, col), keys in self.points.items():
                if name == sheet and top <= row <= bottom and left <= col <= right:
                    yield from keys
        for (name, t, lft, b, r), key in self.areas:
            if name == sheet and t <= bottom and top <= b and lft <= right and left <= r:
                yield key


@dataclass
class Precedents:
    rects: list[Rect]
    unknown: list[CellKey]


def precedents(
    engine: Engine, starts: Iterable[CellKey], *, all_levels: bool, max_formulas: int
) -> Precedents:
    """What the formulas at ``starts`` use; with ``all_levels``, what those use in turn."""
    result = Precedents([], [])
    queue = deque(starts)
    visited = set(queue)
    while queue:
        title, row, col = queue.popleft()
        sheet = engine.sheets[title.casefold()]
        cell = sheet._cells.get((row, col))
        if cell is None or cell.data_type != "f":
            continue
        try:
            found = uses(engine, sheet, row, col)
        except (UncalculableError, FormulaError):
            result.unknown.append((title, row, col))
            continue
        if found.unresolved:
            result.unknown.append((title, row, col))
        result.rects.extend(found.rects)
        if not all_levels:
            continue
        for rect in found.rects:
            for key in _formulas_in(engine, rect):
                if key not in visited:
                    if len(visited) >= max_formulas:
                        raise LimitExceededError(
                            f"Followed more than {max_formulas:,} formulas. Ask for direct "
                            "precedents, or raise --max-cells."
                        )
                    visited.add(key)
                    queue.append(key)
    return result


def _formulas_in(engine: Engine, rect: Rect) -> Iterator[CellKey]:
    title, top, left, bottom, right = rect
    sheet = engine.sheets[title.casefold()]
    last_row, last_col = extent(sheet)
    for (row, col), cell in stored_in(
        sheet, top, left, min(bottom, last_row), min(right, last_col)
    ):
        if cell.data_type == "f":
            yield (title, row, col)
