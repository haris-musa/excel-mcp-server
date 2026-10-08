"""Regenerates tests/fixtures/formula_golden.json from real Microsoft Excel.

Windows with Excel installed only; CI never runs this. It writes the cases of
golden_cases.py into a workbook with the server's own write path, opens that file in
Excel, recalculates everything and stores Excel's results. tests/test_formula_golden.py
then checks the calculator against the stored results without Excel.

    uvx --with pywin32 python scripts/excel_golden.py
"""

import json
import os
import sys
import tempfile
from pathlib import Path
from typing import Any

import win32com.client  # pyright: ignore[reportMissingImports]
from golden_cases import CASES, HIDDEN_ROWS, INPUTS, NAMES, TABLES
from golden_workbook import build_workbook, calculate_cases, case_row, same

from excel_mcp.errors import InvalidFormulaError, UnsafeFormulaError
from excel_mcp.formulas import check_formula

FIXTURE = Path(__file__).resolve().parent.parent / "tests" / "fixtures" / "formula_golden.json"
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


def start_excel() -> Any:
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    return excel


def excel_values(formulas: list[str], directory: Path) -> list[object] | None:
    """Excel's result for each formula, or None if Excel cannot open the workbook.

    Each call starts Excel afresh: after a workbook it cannot open, it often fails to calculate.
    """
    excel = start_excel()
    try:
        return _calculate(excel, formulas, directory)
    finally:
        excel.Quit()


def _calculate(excel: Any, formulas: list[str], directory: Path) -> list[object] | None:
    path = directory / "golden.xlsx"
    build_workbook(path, INPUTS, NAMES, formulas, HIDDEN_ROWS, TABLES)
    temp = Path(os.environ["TEMP"])
    before = set(temp.glob("error*.xml"))
    try:
        workbook = excel.Workbooks.Open(  # pyright: ignore[reportAttributeAccessIssue]
            Filename=str(path), UpdateLinks=0, ReadOnly=True, CorruptLoad=0
        )
    except Exception:
        return None
    try:
        created = set(temp.glob("error*.xml")) - before
        if not created:
            print("no repair log")
        if created:
            for log in created:
                print(log.read_text(errors="replace"))
            raise SystemExit("Excel repaired the workbook; see the log above")
        excel.CalculateFull()  # pyright: ignore[reportAttributeAccessIssue]
        sheet = workbook.Worksheets("Cases")
        values = []
        for index in range(len(formulas)):
            value = sheet.Cells(case_row(index), 1).Value2
            values.append(ERROR_CODES.get(value, value) if isinstance(value, int) else value)
        return values
    finally:
        workbook.Close(False)


def rejected_by_excel(formulas: list[str], directory: Path) -> list[str]:
    """The formulas that stop Excel from opening the workbook, found by halving."""
    if excel_values(formulas, directory) is not None:
        return []
    if len(formulas) == 1:
        return formulas
    middle = len(formulas) // 2
    return [
        *rejected_by_excel(formulas[:middle], directory),
        *rejected_by_excel(formulas[middle:], directory),
    ]


def _allowed(formula: str) -> bool:
    try:
        check_formula(formula, INPUTS)
    except (InvalidFormulaError, UnsafeFormulaError) as error:
        print(f"not written by the server ({error}): {formula}")
        return False
    return True


def main() -> None:
    formulas = [case if isinstance(case, str) else case[0] for case in CASES]
    formulas = [f for f in formulas if _allowed(f)]
    tolerances = {c[0]: c[1] for c in CASES if not isinstance(c, str)}
    with tempfile.TemporaryDirectory() as temp:
        unique = list(dict.fromkeys(formulas))
        rejected = rejected_by_excel(unique, Path(temp))
        for formula in rejected:
            print("rejected by Excel:", formula)
        accepted = [f for f in unique if f not in rejected]
        values = excel_values(accepted, Path(temp))
    assert values is not None
    results = dict(zip(accepted, values, strict=True))
    excel = start_excel()
    try:
        version = excel.Version
    finally:
        excel.Quit()
    cases = [
        {
            "formula": f,
            "expected": results[f],
            **({"tolerance": tolerances[f]} if f in tolerances else {}),
        }
        for f in accepted
    ]
    with tempfile.TemporaryDirectory() as temp:
        path = Path(temp) / "golden.xlsx"
        build_workbook(path, INPUTS, NAMES, accepted, HIDDEN_ROWS, TABLES)
        actual = calculate_cases(path, len(accepted))
    uncalculated = [i for i, value in enumerate(actual) if value is None]
    wrong = [
        (c["formula"], c["expected"], a)
        for c, a in zip(cases, actual, strict=True)
        if a is not None and not same(c["expected"], a, c.get("tolerance", 1e-12))
    ]
    for formula, expected, got in wrong:
        print(f"MISMATCH {formula}: Excel {expected!r}, calculator {got!r}")
    total = len(cases)
    print(
        f"{total} cases, {total - len(uncalculated) - len(wrong)} match, {len(wrong)} mismatch, "
        f"{len(uncalculated)} uncalculated"
    )
    for i in uncalculated:
        print("UNCALCULATED", cases[i]["formula"])
    document = {
        "excel": version,
        "inputs": INPUTS,
        "names": NAMES,
        "hidden_rows": HIDDEN_ROWS,
        "tables": TABLES,
        "cases": cases,
        "uncalculated": uncalculated,
    }
    FIXTURE.write_bytes((json.dumps(document, indent=1, ensure_ascii=False) + "\n").encode("utf-8"))
    sys.exit(1 if wrong else 0)


if __name__ == "__main__":
    main()
