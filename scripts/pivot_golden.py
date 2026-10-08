"""Regenerates tests/fixtures/pivot_golden.json from real Microsoft Excel.

Windows with Excel installed only; CI never runs this. It creates the PivotTables of
pivot_cases.py with the server, opens the workbook in Excel, refreshes every PivotTable and
stores the cells Excel shows. tests/test_pivot_golden.py then checks that the cells the
server writes are the ones Excel shows after a refresh, without Excel. Cases where the two
differ are listed and left out of the fixture.

    uvx --with pywin32 python scripts/pivot_golden.py
"""

import json
import os
import sys
import tempfile
from pathlib import Path
from typing import Any

import win32com.client  # pyright: ignore[reportMissingImports]
from pivot_cases import SOURCE, cases
from pivot_workbook import build_workbook, same, stored_grids

FIXTURE = Path(__file__).resolve().parent.parent / "tests" / "fixtures" / "pivot_golden.json"
ERROR_CODES = {
    -2146826281: "#DIV/0!",
    -2146826246: "#N/A",
    -2146826259: "#NAME?",
    -2146826288: "#NULL!",
    -2146826252: "#NUM!",
    -2146826265: "#REF!",
    -2146826273: "#VALUE!",
}


def refreshed_grids(excel: Any, path: Path, names: list[str]) -> dict[str, list[list[Any]]]:
    temp = Path(os.environ["TEMP"])
    before = set(temp.glob("error*.xml"))
    workbook = excel.Workbooks.Open(Filename=str(path), UpdateLinks=0, ReadOnly=True, CorruptLoad=0)
    try:
        created = set(temp.glob("error*.xml")) - before
        if created:
            for log in created:
                print(log.read_text(errors="replace"))
            raise SystemExit("Excel repaired the workbook; see the log above")
        print("no repair log")
        grids = {}
        for name in names:
            table = workbook.Worksheets(name).PivotTables(name)
            table.PivotCache().Refresh()
            data = table.TableRange2.Value2
            grids[name] = [
                [ERROR_CODES.get(cell, cell) if isinstance(cell, int) else cell for cell in row]
                for row in data
            ]
        return grids
    finally:
        workbook.Close(False)


def main() -> None:
    chosen = cases()
    names = [case["name"] for case in chosen]
    with tempfile.TemporaryDirectory() as directory:
        path = Path(directory) / "pivots.xlsx"
        workbook = build_workbook(chosen)
        ours = stored_grids(workbook)
        workbook.save(path)
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        try:
            theirs = refreshed_grids(excel, path, names)
        finally:
            excel.Quit()
    kept, differing = [], []
    for case in chosen:
        expected, actual = theirs[case["name"]], ours[case["name"]]
        equal = len(expected) == len(actual) and all(
            len(left) == len(right) and all(same(a, b) for a, b in zip(left, right, strict=True))
            for left, right in zip(expected, actual, strict=True)
        )
        (kept if equal else differing).append({**case, "expected": expected})
    for case in differing:
        print("DIFFERS", {key: value for key, value in case.items() if key != "expected"})
    FIXTURE.write_text(
        json.dumps({"source": SOURCE, "cases": kept}, default=str, separators=(",", ":")) + "\n",
        encoding="utf-8",
        newline="\n",
    )
    print(f"{len(kept)} cases stored, {len(differing)} differ")
    sys.exit(1 if differing else 0)


if __name__ == "__main__":
    main()
