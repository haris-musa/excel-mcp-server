"""Checks the cells written for PivotTables against what Excel shows after a refresh.

The fixture is written by scripts/pivot_golden.py (Windows and Excel only); this test
needs neither.
"""

import json
from pathlib import Path

from pivot_cases import SOURCE
from pivot_workbook import build_workbook, same, stored_grids

FIXTURE = Path(__file__).parent / "fixtures" / "pivot_golden.json"


def test_pivot_cells_match_excel() -> None:
    golden = json.loads(FIXTURE.read_text(encoding="utf-8"))
    assert golden["source"] == json.loads(json.dumps(SOURCE, default=str)), "regenerate the fixture"
    cases = golden["cases"]
    grids = stored_grids(
        build_workbook(
            [{key: value for key, value in case.items() if key != "expected"} for case in cases]
        )
    )

    wrong = []
    for case in cases:
        actual = grids[case["name"]]
        expected = case["expected"]
        equal = len(actual) == len(expected) and all(
            len(left) == len(right) and all(same(a, b) for a, b in zip(left, right, strict=True))
            for left, right in zip(expected, actual, strict=True)
        )
        if not equal:
            wrong.append(
                f"{case['name']}: { ({k: v for k, v in case.items() if k != 'expected'}) }"
            )
    assert not wrong, "\n".join(wrong)
    assert len(cases) >= 200
