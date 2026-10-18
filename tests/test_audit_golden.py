"""Checks formula tracing, Evaluate Formula and Goal Seek against results recorded from Excel.

The fixture is written by scripts/audit_golden.py (Windows and Excel only); this test needs
neither. Excel reports precedents and dependents on the same sheet only, so those are compared
for the cells of 'Sales'.
"""

import json
import shutil
from pathlib import Path
from typing import Any

import pytest
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

from excel_mcp.config import Limits
from excel_mcp.errors import ExcelMCPError
from excel_mcp.operations import goal_seek as seeking
from excel_mcp.operations import trace as tracing
from excel_mcp.paths import PathPolicy
from excel_mcp.workspace import Workspace

FIXTURE = Path(__file__).parent / "fixtures" / "audit_golden.json"
MAX_CHANGE = 0.001


@pytest.fixture(scope="module")
def golden() -> dict[str, Any]:
    return json.loads(FIXTURE.read_text(encoding="utf-8"))


@pytest.fixture(scope="module")
def folder(tmp_path_factory: pytest.TempPathFactory) -> Path:
    folder = tmp_path_factory.mktemp("audit")
    build_workbook(folder / FILE)
    return folder


@pytest.fixture(scope="module")
def workspace(folder: Path) -> Workspace:
    return Workspace(PathPolicy([folder]), Limits(), allow_macro_workbooks=False)


def test_the_fixture_is_for_these_cases(golden: dict[str, Any]) -> None:
    traced = [(c["range"], c["all_levels"]) for c in golden["trace"]]
    assert traced == [*PRECEDENTS, *DEPENDENTS]
    assert [c["cell"] for c in golden["evaluate"]] == EVALUATE
    assert [c["name"] for c in golden["goal_seek"]] == [c["name"] for c in GOALS]


def test_precedents_and_dependents_match_excel(
    workspace: Workspace, golden: dict[str, Any]
) -> None:
    wrong = []
    for case in golden["trace"]:
        found = tracing.trace_formula(
            workspace,
            FILE,
            "Sales",
            case["range"],
            case["direction"],
            all_levels=case["all_levels"],
        )
        ours = found.cells.get("Sales", [])
        if cell_set(ours) != cell_set(case["expected"].split(",")):
            wrong.append(
                f"{case['direction']} {case['range']}: Excel {case['expected']}, ours {ours}"
            )
    assert not wrong, "\n".join(wrong)
    assert len(golden["trace"]) >= 30


def test_evaluate_steps_are_the_ones_checked_against_excel(
    workspace: Workspace, golden: dict[str, Any]
) -> None:
    for case in golden["evaluate"]:
        found = tracing.trace_formula(
            workspace, FILE, "Sales", case["cell"], "evaluate", all_levels=False
        )
        steps = [{"expr": step.expr, "value": step.value} for step in found.steps]
        assert steps == case["steps"], case["cell"]


def test_goal_seek_matches_excel(
    workspace: Workspace, golden: dict[str, Any], folder: Path
) -> None:
    for index, (case, recorded) in enumerate(zip(GOALS, golden["goal_seek"], strict=True)):
        copy = f"seek{index}.xlsx"
        shutil.copy(folder / FILE, folder / copy)
        try:
            ours = seeking.goal_seek(
                workspace,
                copy,
                goal_sheet(index),
                goal_cell(case),
                case["target"],
                goal_changing(case),
            )
        except ExcelMCPError:
            assert not recorded["found"], case["name"]  # Excel found nothing either
            continue
        if ours.converged:
            assert abs(ours.result - case["target"]) < MAX_CHANGE, case["name"]
        if recorded["found"]:
            assert ours.converged, case["name"]
            slack = 2e-3 * max(1.0, abs(recorded["x"]))
            assert ours.value == pytest.approx(recorded["x"], abs=slack), case["name"]
            assert abs(recorded["f"] - case["target"]) < MAX_CHANGE
        else:
            # Excel gave up; this search may still find a root (a cubic from a flat start).
            assert ours.iterations > 0
    assert sum(r["found"] for r in golden["goal_seek"]) >= 15
