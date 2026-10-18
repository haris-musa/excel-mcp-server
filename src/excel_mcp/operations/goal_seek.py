"""Goal Seek: change one cell until a formula elsewhere reaches a value."""

from openpyxl.cell.cell import Cell, MergedCell
from openpyxl.worksheet.worksheet import Worksheet
from pydantic import BaseModel

from excel_mcp.calc.cells import CellKey, constant
from excel_mcp.calc.dependencies import Dependents, Rect, precedents
from excel_mcp.calc.engine import Engine
from excel_mcp.calc.solver import seek
from excel_mcp.calc.steps import shown
from excel_mcp.calc.values import Scalar, UncalculableError
from excel_mcp.errors import InvalidArgumentError
from excel_mcp.refs import CellRange, parse_cell
from excel_mcp.workspace import Workspace, get_sheet


class GoalSeekResult(BaseModel):
    converged: bool
    value: float
    result: float
    iterations: int
    note: str | None = None


def goal_seek(
    workspace: Workspace, path: str, sheet: str, cell: str, target: float, changing_cell: str
) -> GoalSeekResult:
    """Find the value for ``changing_cell`` that makes the formula in ``cell`` equal ``target``,
    and store it there when the search converged."""
    with workspace.read_lazy(path) as (stored, formulas):
        engine = Engine(stored, formulas, workspace.display(workspace.resolve(path)))
        problem = _Problem(engine, workspace, get_sheet(formulas, sheet), cell, changing_cell)
        outcome = seek(problem.evaluate, problem.start, target)
        if outcome is None:
            raise InvalidArgumentError(
                f"The formula in {cell} gives {shown(problem.last)} at the current value of "
                f"{changing_cell}; it must give a number."
            )
    best = outcome.best
    value = float(f"{best.x:.15g}")  # the 15 digits Excel shows
    result = GoalSeekResult(
        converged=outcome.reached, value=value, result=best.value, iterations=outcome.iterations
    )
    if not outcome.reached:
        result.note = (
            f"Goal Seeking with Cell {cell} may not have found a solution. "
            f"{changing_cell} was not changed."
        )
        return result
    with workspace.edit(path) as workbook:
        get_sheet(workbook, sheet)[changing_cell].value = value
    return result


class _Problem:
    """The input cell to vary and the formula cell to read, calculated by the engine."""

    def __init__(
        self, engine: Engine, workspace: Workspace, sheet: Worksheet, goal: str, changing: str
    ) -> None:
        self.engine = engine
        self.sheet = sheet
        self.goal = parse_cell(goal)
        self.name = goal
        self.last: Scalar = None
        self.calls = 0
        formula = sheet._cells.get(self.goal)
        if formula is None or formula.data_type != "f":
            raise InvalidArgumentError(f"{goal} holds no formula to set to a value.")
        self.changing, self.start = _input(sheet, *parse_cell(changing), changing)
        engine.stale = _stale(engine, workspace, sheet, self.goal, goal, changing)

    def evaluate(self, x: float) -> float | None:
        """The formula's number with the input set to ``x``; None for text or an error value."""
        self.changing.value = x
        for key in self.engine.stale:
            self.engine.memo.pop(key, None)
            self.engine.spills.pop(key, None)
        self.engine.work = 0
        self.calls += 1
        try:
            self.last = self.engine.calculate(self.sheet, *self.goal)
        except UncalculableError as reason:
            if self.calls > 1:
                return None  # too large or overflowing at this x; the search tries another
            raise InvalidArgumentError(
                f"The formula in {self.name} cannot be calculated here: {reason}."
            ) from None
        if isinstance(self.last, bool) or not isinstance(self.last, int | float):
            return None
        return float(self.last)


def _input(sheet: Worksheet, row: int, col: int, name: str) -> tuple[Cell, float]:
    """The cell Goal Seek changes and the number it holds (a blank counts as 0)."""
    cell = sheet.cell(row, col)
    if isinstance(cell, MergedCell):
        raise InvalidArgumentError(f"{name} is inside a merged range; use its top-left cell.")
    if cell.data_type == "f":
        raise InvalidArgumentError(
            f"{name} holds a formula; Goal Seek changes a constant. Pick an input cell."
        )
    value = constant(cell)
    if value is None:
        return cell, 0.0
    if isinstance(value, bool) or not isinstance(value, int | float):
        raise InvalidArgumentError(f"{name} holds {shown(value)}, not a number.")
    return cell, float(value)


def _stale(
    engine: Engine,
    workspace: Workspace,
    sheet: Worksheet,
    goal: tuple[int, int],
    goal_name: str,
    changing_name: str,
) -> set[CellKey]:
    """The formula cells whose stored results no longer hold once the input changes."""
    limit = workspace.limits.max_cells
    index = Dependents(engine, limit)
    row, col = parse_cell(changing_name)
    seeds = [(sheet.title, row, col, row, col)]
    seeds += [(title, r, c, r, c) for title, r, c in index.unknown]
    chain = index.of(seeds, all_levels=True, limit=limit)
    goal_key = (sheet.title, *goal)
    if goal_key not in chain:
        raise InvalidArgumentError(f"{goal_name} does not depend on {changing_name}.")
    reads = precedents(engine, [goal_key], all_levels=True, max_formulas=limit).rects
    if any(_reads_body(key, index.arrays[key], reads) for key in chain if key in index.arrays):
        raise InvalidArgumentError(
            f"{goal_name} reads cells filled by a spilling formula, which cannot be recalculated "
            "here. Refer to the spilled range with # instead."
        )
    return set(chain) | set(index.unknown)


def _reads_body(anchor: CellKey, spill: Rect, reads: list[Rect]) -> bool:
    """Whether any of the rectangles covers a cell the formula at ``anchor`` spills into."""
    body = CellRange(*spill[1:]).minus(CellRange(anchor[1], anchor[2], anchor[1], anchor[2]))
    read_areas = [CellRange(*read[1:]) for read in reads if read[0] == spill[0]]
    return any(part.overlaps(area) for area in read_areas for part in body)
