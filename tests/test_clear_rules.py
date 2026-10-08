"""Clearing conditional formats and data validation with clear_range, as Excel does."""

import re
from pathlib import Path

import pytest
from openpyxl import load_workbook

from tests.conftest import ToolCall
from tests.package_support import (
    assert_package_is_consistent,
    copy_fixture,
    read_parts,
    sheet_part,
    text,
)

pytestmark = pytest.mark.anyio

SHEET = {"path": "sales.xlsx", "sheet": "Data"}
VALIDATION = {"type": "whole", "operator": "between", "minimum": "1", "maximum": "9"}


async def _rules(call: ToolCall, range: str = "A1:A10") -> None:
    rule = {"type": "formula", "formula": f"={range.split(':')[0]}>5", "fill_color": "#FFC7CE"}
    await call("add_conditional_format", **SHEET, range=range, rule=rule)
    await call("add_data_validation", **SHEET, range=range, rule=VALIDATION)


def _state(path: Path) -> tuple[list[tuple[str, list[str]]], list[str]]:
    sheet = load_workbook(path)["Data"]
    formats = [
        (str(entry.sqref), [rule.formula[0] for rule in entry.rules])
        for entry in sheet.conditional_formatting
    ]
    return formats, [str(v.sqref) for v in sheet.data_validations.dataValidation]


def _x14(path: Path) -> str:
    parts = read_parts(path)
    return text(parts, sheet_part(parts, "Data"))


async def test_a_rule_keeps_the_cells_that_are_not_cleared(call: ToolCall, sample: Path) -> None:
    await _rules(call)
    await call("clear_range", **SHEET, range="A3:A4", clear="rules")
    assert _state(sample) == ([("A1:A2 A5:A10", ["A1>5"])], ["A1:A2 A5:A10"])


async def test_formulas_follow_the_new_top_left_cell(call: ToolCall, sample: Path) -> None:
    await _rules(call)
    await call("clear_range", **SHEET, range="A1:A2", clear="rules")
    assert _state(sample) == ([("A3:A10", ["A3>5"])], ["A3:A10"])


async def test_rules_cleared_from_all_their_cells_are_deleted(call: ToolCall, sample: Path) -> None:
    await _rules(call)
    await call("clear_range", **SHEET, range="A1:XFD1048576", clear="rules")
    assert _state(sample) == ([], [])


@pytest.mark.parametrize(
    ("clear", "formats", "validation"),
    [("contents", True, True), ("formats", False, True), ("rules", False, False)],
)
async def test_what_each_kind_of_clearing_removes(
    call: ToolCall, sample: Path, clear: str, formats: bool, validation: bool
) -> None:
    await _rules(call, "C2:C5")
    await call("clear_range", **SHEET, range="C2:C5", clear=clear)
    found_formats, found_validation = _state(sample)
    assert bool(found_formats) is formats
    assert bool(found_validation) is validation


async def test_clearing_all_removes_rules_too(call: ToolCall, sample: Path) -> None:
    await _rules(call, "C2:C5")
    await call("clear_range", **SHEET, range="C2:C3", clear="all")
    assert _state(sample) == ([("C4:C5", ["C4>5"])], ["C4:C5"])


async def test_rules_elsewhere_are_untouched(call: ToolCall, sample: Path) -> None:
    await _rules(call, "C2:C5")
    await call("clear_range", **SHEET, range="A1:B20", clear="rules")
    assert _state(sample) == ([("C2:C5", ["C2>5"])], ["C2:C5"])


async def test_extended_rules_lose_cells_with_their_classic_partner(
    call: ToolCall, sample: Path
) -> None:
    bar = {"type": "data_bar", "colors": ["#638EC6"]}
    await call("add_conditional_format", **SHEET, range="C2:C5", rule=bar)
    await call(
        "add_conditional_format",
        **SHEET,
        range="D2:D5",
        rule={"type": "icon_set", "icon_set": "3Stars"},
    )
    await call("clear_range", **SHEET, range="C4:D4", clear="rules")
    xml = _x14(sample)
    assert re.findall(r"<xm:sqref>([^<]*)</xm:sqref>", xml) == ["C2:C3 C5", "D2:D3 D5"]
    assert re.search(r'<cfRule type="dataBar" priority="1">', xml)
    assert_package_is_consistent(read_parts(sample))

    await call("clear_range", **SHEET, range="C1:D9", clear="rules")
    xml = _x14(sample)
    assert "x14:cfRule" not in xml
    assert "dataBar" not in xml
    assert_package_is_consistent(read_parts(sample))


async def test_what_excel_saved_can_be_cleared(call: ToolCall, files: Path) -> None:
    saved = copy_fixture(files, "excel_sparklines.xlsx", "excel.xlsx")
    part = sheet_part(read_parts(saved), "Data")
    assert b"x14:cfRule" in read_parts(saved)[part]
    await call("clear_range", path="excel.xlsx", sheet="Data", range="B2:C4", clear="rules")
    kept = read_parts(saved)[part]
    assert b"<xm:sqref>B5:B7</xm:sqref>" in kept
    await call("clear_range", path="excel.xlsx", sheet="Data", range="A1:XFD1048576", clear="rules")
    parts = read_parts(saved)
    for marker in (b"x14:cfRule", b"<cfRule", b"<dataValidation", b"x14:conditionalFormatting"):
        assert marker not in parts[part]
    assert b"x14:sparkline" in parts[part]
    assert_package_is_consistent(parts)
