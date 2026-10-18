"""Formula auditing: what a formula uses, what uses a cell, and how a formula evaluates."""

from collections import defaultdict
from collections.abc import Iterable
from typing import Literal

from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet
from pydantic import BaseModel, Field

from excel_mcp.calc.cells import CellKey, stored_in
from excel_mcp.calc.dependencies import Dependents, Rect, precedents
from excel_mcp.calc.engine import Engine
from excel_mcp.calc.steps import evaluate, sheet_text, shown
from excel_mcp.calc.values import UncalculableError
from excel_mcp.errors import InvalidArgumentError
from excel_mcp.refs import MAX_COLUMN, MAX_ROW, CellRange, parse_range
from excel_mcp.workspace import Workspace, get_sheet

Direction = Literal["precedents", "dependents", "evaluate"]
Box = tuple[int, int, int, int]  # top, left, bottom, right
MAX_LISTED = 200
MAX_STEPS = 100


class Step(BaseModel):
    expr: str
    value: str


class Trace(BaseModel):
    cells: dict[str, list[str]] = Field(default_factory=dict)
    unknown: list[str] = []
    steps: list[Step] = []
    more: int = 0


def trace_formula(
    workspace: Workspace,
    path: str,
    sheet: str,
    range: str,
    direction: Direction,
    *,
    all_levels: bool,
) -> Trace:
    area = parse_range(range).within(workspace.limits.max_cells)
    if direction == "evaluate" and (area.size != 1 or all_levels):
        raise InvalidArgumentError("evaluate takes one cell and no all_levels.")
    limit = workspace.limits.max_cells
    with workspace.read_lazy(path) as (stored, formulas):
        engine = Engine(stored, formulas, workspace.display(workspace.resolve(path)))
        home = get_sheet(formulas, sheet)
        titles = [other.title for other in engine.sheets.values()]
        if direction == "dependents":
            index = Dependents(engine, limit)
            users = index.of([_rect(home, area)], all_levels=all_levels, limit=limit)
            filled = [index.span(key) for key in users]
            boxes = [(t, (top, left, bottom, right)) for t, top, left, bottom, right in filled]
            return _listed(boxes, _cells(index.unknown), titles)
        starts = [
            (home.title, *position)
            for position, cell in stored_in(
                home, area.min_row, area.min_col, area.max_row, area.max_col
            )
            if cell.data_type == "f"
        ]
        if not starts:
            raise InvalidArgumentError(f"{range} holds no formula.")
        if direction == "evaluate":
            return _evaluate(engine, home, area, range)
        found = precedents(engine, starts, all_levels=all_levels, max_formulas=limit)
        boxes = [(t, (top, left, bottom, right)) for t, top, left, bottom, right in found.rects]
        return _listed(boxes, _cells(found.unknown), titles)


def _rect(sheet: Worksheet, area: CellRange) -> Rect:
    return (sheet.title, area.min_row, area.min_col, area.max_row, area.max_col)


def _cells(keys: Iterable[CellKey]) -> list[tuple[str, Box]]:
    return [(title, (row, col, row, col)) for title, row, col in keys]


def _evaluate(engine: Engine, home: Worksheet, area: CellRange, range: str) -> Trace:
    try:
        steps = evaluate(engine, home, area.min_row, area.min_col, MAX_STEPS)
    except UncalculableError as reason:
        raise InvalidArgumentError(f"{range} cannot be evaluated here: {reason}.") from None
    return Trace(
        steps=[Step(expr=text, value=shown(value)) for text, value in steps.shown],
        more=steps.skipped,
    )


def _ranges(group: Iterable[tuple[str, Box]], titles: list[str]) -> list[tuple[str, str]]:
    """(sheet, range) for the cells of ``group``, merged, in the workbook's order of sheets."""
    by_sheet: defaultdict[str, set[Box]] = defaultdict(set)
    for title, box in group:
        by_sheet[title].add(box)
    return [
        (title, _text(box))
        for title in titles
        if title in by_sheet
        for box in merged(by_sheet[title])
    ]


def _listed(
    boxes: Iterable[tuple[str, Box]], unknown: Iterable[tuple[str, Box]], titles: list[str]
) -> Trace:
    found, unsure = _ranges(boxes, titles), _ranges(unknown, titles)
    cells: dict[str, list[str]] = {}
    for title, text in found[:MAX_LISTED]:
        cells.setdefault(title, []).append(text)
    room = max(MAX_LISTED - len(found), 0)
    return Trace(
        cells=cells,
        unknown=[f"{sheet_text(title)}!{text}" for title, text in unsure[:room]],
        more=max(len(found) - MAX_LISTED, 0) + max(len(unsure) - room, 0),
    )


def merged(boxes: set[Box]) -> list[Box]:
    """The smallest list of rectangles covering the same cells, found by joining neighbours."""
    large = sorted((b for b in boxes if b[0] != b[2] or b[1] != b[3]), key=_area, reverse=True)
    kept = {b for b in boxes if not any(_inside(b, big) for big in large if big != b)}
    while True:
        joined = _join(_join(kept, vertical=True), vertical=False)
        if len(joined) == len(kept):
            return sorted(joined)
        kept = joined


def _area(box: Box) -> int:
    return (box[2] - box[0] + 1) * (box[3] - box[1] + 1)


def _inside(box: Box, outer: Box) -> bool:
    return outer[0] <= box[0] and outer[1] <= box[1] and box[2] <= outer[2] and box[3] <= outer[3]


def _join(boxes: set[Box], *, vertical: bool) -> set[Box]:
    """Join boxes that share their columns (or rows) and touch or overlap along the other axis."""
    groups: defaultdict[tuple[int, int], list[tuple[int, int]]] = defaultdict(list)
    for top, left, bottom, right in boxes:
        if vertical:
            groups[(left, right)].append((top, bottom))
        else:
            groups[(top, bottom)].append((left, right))
    joined: set[Box] = set()
    for (low, high), spans in groups.items():
        spans.sort()
        start, end = spans[0]
        for first, last in [*spans[1:], (-1, -1)]:
            if first != -1 and first <= end + 1:
                end = max(end, last)
                continue
            joined.add((start, low, end, high) if vertical else (low, start, high, end))
            start, end = first, last
    return joined


def _text(box: Box) -> str:
    top, left, bottom, right = box
    if top == 1 and bottom == MAX_ROW:
        return f"{get_column_letter(left)}:{get_column_letter(right)}"
    if left == 1 and right == MAX_COLUMN:
        return f"{top}:{bottom}"
    return str(CellRange(top, left, bottom, right))
