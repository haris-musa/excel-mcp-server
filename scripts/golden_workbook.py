"""Builds the golden workbook and runs the calculator on it (shared by the generator and tests)."""

import math
from pathlib import Path
from typing import Any

from openpyxl import Workbook, load_workbook
from openpyxl.workbook.defined_name import DefinedName

from excel_mcp.calc.engine import Engine
from excel_mcp.calc.values import ExcelError, UncalculableError
from excel_mcp.operations.cells import write_range

MAX_CELLS = 1_000_000
DEFAULT_TOLERANCE = 1e-12


def build_workbook(
    path: Path, inputs: dict[str, list[list[Any]]], names: dict[str, str], formulas: list[str]
) -> None:
    """Write the inputs and one formula per row of 'Cases' column A, as the server would."""
    workbook = Workbook()
    workbook.remove(workbook.active)  # pyright: ignore[reportArgumentType]
    for title in inputs:
        workbook.create_sheet(title)
    for title, rows in inputs.items():
        if rows:
            write_range(workbook[title], "A1", rows, MAX_CELLS)  # pyright: ignore[reportArgumentType]
    cases = workbook["Cases"]
    for row, formula in enumerate(formulas, start=1):
        write_range(cases, f"A{row}", [[formula]], MAX_CELLS)  # pyright: ignore[reportArgumentType]
    for name, target in names.items():
        workbook.defined_names[name] = DefinedName(name, attr_text=target)
    workbook.save(path)


def calculate_cases(path: Path, count: int) -> list[object]:
    """The calculator's value for each of the first ``count`` cases, or None if uncalculated."""
    formulas = load_workbook(path)
    cached = load_workbook(path, data_only=True)
    engine = Engine(cached, formulas)
    sheet = formulas["Cases"]
    results: list[object] = []
    for row in range(1, count + 1):
        try:
            value = engine.calculate(sheet, row, 1)
        except UncalculableError:
            results.append(None)
            continue
        results.append(value.code if isinstance(value, ExcelError) else value)
    return results


def same(expected: object, actual: object, tolerance: float = DEFAULT_TOLERANCE) -> bool:
    if isinstance(expected, bool) or isinstance(actual, bool):
        return type(expected) is type(actual) and expected == actual
    if isinstance(expected, int | float) and isinstance(actual, int | float):
        return math.isclose(expected, actual, rel_tol=tolerance, abs_tol=tolerance)
    return expected == actual
