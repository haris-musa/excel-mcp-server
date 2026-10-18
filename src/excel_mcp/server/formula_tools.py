"""Tools for understanding and solving formulas: tracing, step-by-step evaluation, Goal Seek."""

from typing import Annotated

from pydantic import Field

from excel_mcp.operations import goal_seek as seeking
from excel_mcp.operations import trace as tracing
from excel_mcp.operations.goal_seek import GoalSeekResult
from excel_mcp.operations.trace import Direction, Trace
from excel_mcp.server.params import CellRef, RangeRef, SheetName, WorkbookPath
from excel_mcp.server.registry import ToolRegistry
from excel_mcp.workspace import Workspace


def register(tools: ToolRegistry, workspace: Workspace) -> None:
    @tools.reader("Trace formula")
    def trace_formula(
        path: WorkbookPath,
        sheet: SheetName,
        range: Annotated[RangeRef, Field(description="A cell; a range for 'dependents'.")],
        direction: Annotated[
            Direction,
            Field(
                description="'precedents': what the formulas use; 'dependents': formulas using "
                "`range`; 'evaluate': Evaluate Formula steps."
            ),
        ],
        all_levels: Annotated[bool, Field(description="Follow to the end, not one level.")] = False,
    ) -> Trace:
        """Audit formulas, calculated here. `cells`: ranges by sheet (names and table
        references resolved); `unknown`: formulas with references that cannot be worked out (a
        computed INDIRECT); `steps`: innermost first; `more`: results left out."""
        return tracing.trace_formula(
            workspace, path, sheet, range, direction, all_levels=all_levels
        )

    @tools.destroyer("Goal Seek")
    def goal_seek(
        path: WorkbookPath,
        sheet: SheetName,
        cell: Annotated[CellRef, Field(description="Formula cell to set.")],
        target: Annotated[float, Field(description="The value it should reach.")],
        changing_cell: Annotated[CellRef, Field(description="Constant to change.")],
    ) -> GoalSeekResult:
        """Goal Seek: change the constant `changing_cell` until the formula in `cell` equals
        `target`, and store it. Stops within 0.001 after at most 100 iterations; if it does not
        converge, nothing is changed."""
        return seeking.goal_seek(workspace, path, sheet, cell, target, changing_cell)
