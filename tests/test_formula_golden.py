"""Checks the formula calculator against results recorded from real Excel.

The fixture is written by scripts/excel_golden.py (Windows and Excel only); this test
needs neither.
"""

import json
from pathlib import Path

from golden_workbook import DEFAULT_TOLERANCE, build_workbook, calculate_cases, same

FIXTURE = Path(__file__).parent / "fixtures" / "formula_golden.json"


def test_calculator_matches_excel(tmp_path: Path) -> None:
    golden = json.loads(FIXTURE.read_text(encoding="utf-8"))
    cases = golden["cases"]
    path = tmp_path / "golden.xlsx"
    build_workbook(
        path,
        golden["inputs"],
        golden["names"],
        [c["formula"] for c in cases],
        golden["hidden_rows"],
        golden["tables"],
    )

    actual = calculate_cases(path, len(cases))

    wrong = [
        f"{case['formula']}: Excel {case['expected']!r}, calculator {value!r}"
        for case, value in zip(cases, actual, strict=True)
        if value is not None
        and not same(case["expected"], value, case.get("tolerance", DEFAULT_TOLERANCE))
    ]
    assert not wrong, "\n".join(wrong)
    left_out = [i for i, value in enumerate(actual) if value is None]
    assert left_out == golden["uncalculated"]
    assert len(cases) >= 1_000


def test_cells_loaded_on_demand_calculate_the_same(tmp_path: Path) -> None:
    golden = json.loads(FIXTURE.read_text(encoding="utf-8"))
    cases = golden["cases"]
    path = tmp_path / "golden.xlsx"
    build_workbook(
        path,
        golden["inputs"],
        golden["names"],
        [c["formula"] for c in cases],
        golden["hidden_rows"],
        golden["tables"],
    )

    assert calculate_cases(path, len(cases), lazy=True) == calculate_cases(path, len(cases))
