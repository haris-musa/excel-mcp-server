"""Regenerate src/excel_mcp/function_arguments.json by asking Excel how many arguments each
worksheet function takes.

Windows with Microsoft Excel installed only; this is not part of the test suite or CI.
Run it after Excel adds functions:

    uv run --with pywin32 python scripts/probe_function_arguments.py

Every function name this project knows is written into a scratch cell with n copies of
``A1`` as arguments, and Excel's parser accepts or refuses it. A name that Excel accepts
with no arguments and with 254 is a user-defined name, not a function it knows (a
user-defined function takes at most 254), so it is left out. So are LET and LAMBDA, whose
arguments are names, not values, and functions that accept an irregular set of counts.
"""

import json
from itertools import pairwise
from pathlib import Path

import win32com.client
from openpyxl.utils import FORMULAE

from excel_mcp.calc import engine  # noqa: F401  (registers the calculator's functions)
from excel_mcp.calc.registry import FUNCTIONS
from excel_mcp.xlfn import FUTURE_FUNCTIONS

OUTPUT = Path(__file__).parent.parent / "src" / "excel_mcp" / "function_arguments.json"
LARGEST_COUNT = 255
SPECIAL_SYNTAX = {"LET", "LAMBDA"}
EXTRA = ["ANCHORARRAY", "GROUPBY", "PIVOTBY", "PERCENTOF", "TRIMRANGE", "REGEXTEST", "REGEXEXTRACT"]
EXTRA += ["REGEXREPLACE", "ARRAYTOTEXT", "VALUETOTEXT", "IMAGE", "STOCKHISTORY", "TRANSLATE"]


def accepts(sheet, name: str, count: int) -> bool:
    cell = sheet.Range("A20")
    try:
        cell.Formula2 = f"={name}({','.join(['A1'] * count)})"
        return True
    except Exception:
        return False
    finally:
        cell.ClearContents()


def probe(
    sheet, name: str
) -> list[int] | None:  # fewest, most, and 2 when only every other count works
    counts = [n for n in range(21) if accepts(sheet, name, n)]
    if counts and counts[-1] >= 19:
        counts += [n for n in range(counts[-1] + 1, LARGEST_COUNT + 1) if accepts(sheet, name, n)]
    elif accepts(sheet, name, LARGEST_COUNT):
        counts.append(LARGEST_COUNT)
    if not counts or (counts[0] == 0 and counts[-1] >= LARGEST_COUNT - 1):
        return None
    gaps = {later - earlier for earlier, later in pairwise(counts)}
    if not (gaps <= {1} or gaps == {2}):
        return None
    return [counts[0], counts[-1], 2 if gaps == {2} else 1]


def main() -> None:
    names = sorted({*FORMULAE, *FUTURE_FUNCTIONS, *FUNCTIONS, *EXTRA} - SPECIAL_SYNTAX)
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        sheet = excel.Workbooks.Add().Worksheets(1)
        table = {name: found for name in names if (found := probe(sheet, name))}
    finally:
        excel.Quit()
    OUTPUT.write_text(
        json.dumps(table, separators=(",", ":")) + "\n", encoding="utf-8", newline="\n"
    )
    print(f"{len(table)} of {len(names)} names written to {OUTPUT}")


if __name__ == "__main__":
    main()
