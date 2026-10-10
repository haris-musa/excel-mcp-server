"""Orders the calculation of formula cells so that long dependency chains need no recursion."""

import logging
from collections.abc import Iterator
from contextlib import suppress
from typing import TYPE_CHECKING

from openpyxl.worksheet.formula import ArrayFormula
from openpyxl.worksheet.worksheet import Worksheet

from excel_mcp.calc.cells import key, stored_in
from excel_mcp.calc.parser import (
    ArrayLiteral,
    Binary,
    Call,
    Name,
    Node,
    Percent,
    Ref,
    Unary,
    parse,
)
from excel_mcp.calc.values import FormulaError, Scalar, UncalculableError
from excel_mcp.lazy_workbook import extent

if TYPE_CHECKING:
    from excel_mcp.calc.engine import Engine

logger = logging.getLogger(__name__)

MAX_NAME_DEPTH = 8


class TooDeepError(Exception):
    """A chain of dependent cells is too long to follow in one go."""

    def __init__(self, sheet: Worksheet, row: int, col: int) -> None:
        super().__init__("formula chain too deep")
        self.target = (sheet, row, col)


class Scheduler:
    """Calculates the formula cells a cell refers to first, from an explicit stack."""

    def __init__(self, engine: "Engine") -> None:
        self.engine = engine

    def calculate(self, sheet: Worksheet, row: int, col: int) -> Scalar:
        """The value of a cell, calculating its formula if Excel stored no result.

        References the scan cannot see (OFFSET, INDEX) fall back to recursion, which is
        depth-limited.
        """
        engine = self.engine
        stack = [(sheet, row, col)]
        queued = {(sheet.title, row, col)}
        while True:
            current = stack[-1]
            try:
                waiting = [
                    d for d in self.unfinished_dependencies(*current) if key(d) not in queued
                ]
                if waiting:
                    stack.extend(reversed(waiting))
                    queued.update(key(d) for d in waiting)
                    continue
                value = engine.cell_value(*current)
            except TooDeepError as deep:
                stack.append(deep.target)
                queued.add(key(deep.target))
                continue
            except UncalculableError:
                # A failed dependency is remembered; the cells that use it will report it.
                if len(stack) == 1:
                    raise
                value = None
            except Exception as error:
                # A hostile or unusual formula must end as an uncalculated cell, never as a
                # failed read: this covers recursion limits, memory and arithmetic overflow.
                logger.debug("calculation failed: %r", error)
                failure = UncalculableError("not supported")
                engine.memo[key(current)] = failure
                if len(stack) == 1:
                    raise failure from None
                value = None
            stack.pop()
            queued.discard(key(current))
            if not stack:
                return value

    def calculate_precedents(self, tree: Node, sheet: Worksheet) -> None:
        """Calculate the cells a formula uses. A failure is remembered, and the formula
        reports it if it needs the cell."""
        for target, top, left, bottom, right in list(self._references(tree, sheet, 0)):
            self.engine.charge(1)
            for position, other in stored_in(target, top, left, bottom, right):
                if other.data_type == "f" and self._is_unfinished(target, *position):
                    with suppress(UncalculableError):
                        self.calculate(target, *position)

    def unfinished_dependencies(
        self, sheet: Worksheet, row: int, col: int
    ) -> list[tuple[Worksheet, int, int]]:
        """Formula cells that the formula at this position names and that are not done yet."""
        engine = self.engine
        cell = sheet._cells.get((row, col))
        if cell is None or cell.data_type != "f" or isinstance(cell.value, ArrayFormula):
            return []
        if not self._is_unfinished(sheet, row, col):
            return []
        try:
            tree = engine.tree(cell)
            references = list(self._references(tree, sheet, 0))
        except (UncalculableError, FormulaError):
            return []
        found = []
        for target, top, left, bottom, right in references:
            engine.charge(1)
            for position, other in stored_in(target, top, left, bottom, right):
                if other.data_type == "f" and self._is_unfinished(target, *position):
                    found.append((target, *position))
        return found

    def _is_unfinished(self, sheet: Worksheet, row: int, col: int) -> bool:
        return (sheet.title, row, col) not in self.engine.memo and (
            self.engine.cached_value(sheet, row, col) is None
        )

    def _references(
        self, node: Node, home: Worksheet, depth: int
    ) -> Iterator[tuple[Worksheet, int, int, int, int]]:
        """The rectangles a syntax tree refers to, as (sheet, top, left, bottom, right)."""
        engine = self.engine
        match node:
            case Ref():
                for target in engine.sheets_in(node, home):
                    yield (
                        target,
                        node.top or 1,
                        node.left or 1,
                        node.bottom or extent(target)[0],
                        node.right or extent(target)[1],
                    )
            case Name() if depth < MAX_NAME_DEPTH:
                scope = engine.sheets.get((node.sheet or "").casefold(), home)
                try:
                    inner = parse(engine.lookup_name(scope, node.name))
                except UncalculableError:
                    return
                yield from self._references(inner, home, depth + 1)
            case Unary(_, operand) | Percent(operand):
                yield from self._references(operand, home, depth)
            case Binary(_, left, right):
                yield from self._references(left, home, depth)
                yield from self._references(right, home, depth)
            case Call():
                for argument in node.args:
                    yield from self._references(argument, home, depth)
            case ArrayLiteral(rows):
                for row in rows:
                    for item in row:
                        yield from self._references(item, home, depth)
