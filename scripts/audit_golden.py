"""Regenerates tests/fixtures/audit_golden.json from real Microsoft Excel.

Windows with Excel installed only; CI never runs this. It builds the workbook of audit_cases.py
with the server's own write path, opens it in Excel and records:

- the cells Excel reports for `Range.DirectPrecedents`, `Precedents`, `DirectDependents` and
  `Dependents` (Excel reports cells of the same sheet only),
- the steps of `trace_formula` evaluate, each checked against `Worksheet.Evaluate` of the
  sub-expression, and
- the result of `Range.GoalSeek` for the cases of audit_cases.py.

tests/test_audit_golden.py then checks the server against the record, without Excel. Anything
where the server disagrees with Excel is listed and the script exits with an error.

    uv run --with pywin32 python scripts/audit_golden.py
"""

import json
import os
import shutil
import sys
import tempfile
from pathlib import Path
from typing import Any

import win32com.client  # pyright: ignore[reportMissingImports]
from audit_cases import (
    DEPENDENTS,
    EVALUATE,
    FILE,
    GOALS,
    PRECEDENTS,
    build_workbook,
    goal_cell,
    goal_changing,
    goal_sheet,
)
from audit_check import cell_set

from excel_mcp.calc.steps import scalar_text
from excel_mcp.calc.values import ExcelError
from excel_mcp.config import Limits
from excel_mcp.errors import ExcelMCPError
from excel_mcp.operations import goal_seek as seeking
from excel_mcp.operations import trace as tracing
from excel_mcp.paths import PathPolicy
from excel_mcp.workspace import Workspace

FIXTURE = Path(__file__).resolve().parent.parent / "tests" / "fixtures" / "audit_golden.json"
ERROR_CODES = {
    -2146826281: "#DIV/0!",
    -2146826246: "#N/A",
    -2146826259: "#NAME?",
    -2146826288: "#NULL!",
    -2146826252: "#NUM!",
    -2146826265: "#REF!",
    -2146826273: "#VALUE!",
    -2146826238: "#CALC!",
}


def excel_address(cell: Any, direction: str, all_levels: bool) -> str:
    name = ("Precedents" if all_levels else "DirectPrecedents") if direction == "precedents" else (
        "Dependents" if all_levels else "DirectDependents"
    )  # fmt: skip
    try:
        return getattr(cell, name).Address
    except Exception:  # Excel raises when there are none on the sheet
        return ""


def text_of(raw: Any) -> str:
    """A value from Excel as `trace_formula` shows values."""
    is_error = isinstance(raw, int) and not isinstance(raw, bool)
    return scalar_text(ExcelError(ERROR_CODES[raw]) if is_error else raw)


def evaluated(sheet: Any, expression: str) -> Any:
    """`Worksheet.Evaluate`, which returns a Range for a reference: its value."""
    result = sheet.Evaluate(expression)
    plain = result is None or isinstance(result, bool | int | float | str | tuple)
    return result if plain else result.Value2


def main() -> None:
    differing: list[str] = []
    with tempfile.TemporaryDirectory() as directory:
        root = Path(directory)
        path = root / FILE
        build_workbook(path)
        workspace = Workspace(PathPolicy([root]), Limits(), allow_macro_workbooks=False)
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        temp = Path(os.environ["TEMP"])
        before = set(temp.glob("error*.xml"))
        try:
            book = excel.Workbooks.Open(Filename=str(path), UpdateLinks=0, CorruptLoad=0)
            created = set(temp.glob("error*.xml")) - before
            if created:
                for log in created:
                    print(log.read_text(errors="replace"))
                raise SystemExit("Excel repaired the workbook; see the log above")
            print("no repair log")
            excel.CalculateFull()
            sales = book.Worksheets("Sales")
            record: dict[str, Any] = {"excel": excel.Version}
            record["trace"] = trace_cases(workspace, sales, differing)
            record["evaluate"] = evaluate_cases(workspace, sales, differing)
            record["goal_seek"] = goal_cases(excel, book, workspace, root, differing)
            book.Close(False)
        finally:
            excel.Quit()
    FIXTURE.write_text(json.dumps(record, indent=1) + "\n", encoding="utf-8", newline="\n")
    print("\n".join(differing) or "all agree")
    sys.exit(1 if differing else 0)


def trace_cases(workspace: Workspace, sales: Any, differing: list[str]) -> list[dict[str, Any]]:
    cases = []
    for direction, items in (("precedents", PRECEDENTS), ("dependents", DEPENDENTS)):
        for range_, all_levels in items:
            expected = excel_address(sales.Range(range_), direction, all_levels)
            ours = tracing.trace_formula(
                workspace,
                FILE,
                "Sales",
                range_,
                direction,
                all_levels=all_levels,  # pyright: ignore[reportArgumentType]
            )
            actual = ours.cells.get("Sales", [])
            if cell_set(expected.split(",")) != cell_set(actual):
                differing.append(
                    f"{direction} {range_} all={all_levels}: Excel {expected}, ours {actual}"
                )
            cases.append(
                {
                    "range": range_,
                    "direction": direction,
                    "all_levels": all_levels,
                    "expected": expected,
                }
            )
    return cases


def evaluate_cases(workspace: Workspace, sales: Any, differing: list[str]) -> list[dict[str, Any]]:
    cases = []
    for cell in EVALUATE:
        try:
            steps = tracing.trace_formula(
                workspace, FILE, "Sales", cell, "evaluate", all_levels=False
            ).steps
        except ExcelMCPError as error:
            differing.append(f"evaluate {cell}: {error}")
            continue
        final = sales.Range(cell).Value2
        checked = []
        for step in steps:
            expected = evaluated(sales, step.expr)
            is_array = isinstance(expected, tuple) or step.value.startswith("{")
            if not is_array and text_of(expected) != step.value:
                excel = text_of(expected)
                differing.append(
                    f"evaluate {cell} step {step.expr}: Excel {excel}, ours {step.value}"
                )
            checked.append({"expr": step.expr, "value": step.value})
        if not isinstance(final, tuple) and text_of(final) != steps[-1].value:
            differing.append(
                f"evaluate {cell}: last step {steps[-1].value}, Excel {text_of(final)}"
            )
        cases.append({"cell": cell, "steps": checked})
    return cases


def goal_cases(
    excel: Any, book: Any, workspace: Workspace, root: Path, differing: list[str]
) -> list[dict[str, Any]]:
    cases = []
    for index, case in enumerate(GOALS):
        sheet = book.Worksheets(goal_sheet(index))
        changing, goal = goal_changing(case), goal_cell(case)
        try:
            found = bool(
                sheet.Range(goal).GoalSeek(Goal=case["target"], ChangingCell=sheet.Range(changing))
            )
        except Exception as error:
            found = False
            print(f"{case['name']}: Excel raised {error}")
        excel.Calculate()
        record = {
            "name": case["name"],
            "found": found,
            "x": sheet.Range(changing).Value2,
            "f": sheet.Range(goal).Value2,
        }
        cases.append(record)
        copy = root / f"seek{index}.xlsx"
        shutil.copy(root / FILE, copy)
        try:
            ours = seeking.goal_seek(
                workspace, copy.name, goal_sheet(index), goal, case["target"], changing
            )
        except ExcelMCPError as error:
            print(f"{case['name']}: Excel found={found} x={record['x']}; ours refuses: {error}")
            continue
        print(
            f"{case['name']}: Excel found={found} x={record['x']} f={record['f']}; "
            f"ours converged={ours.converged} x={ours.value} f={ours.result} n={ours.iterations}"
        )
    return cases


if __name__ == "__main__":
    main()
