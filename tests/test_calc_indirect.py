"""INDIRECT, CELL and INFO in the calculator (Excel's own answers are in the golden fixture)."""

from pathlib import Path

import pytest
from openpyxl import Workbook

from excel_mcp.calc.r1c1 import to_a1
from excel_mcp.calc.values import FormulaError
from tests.conftest import ToolCall
from tests.test_calc import read_values

pytestmark = pytest.mark.anyio


@pytest.mark.parametrize(
    ("text", "expected"),
    [
        ("R2C3", "C2"),
        ("r2c3", "C2"),
        ("R[1]C[-1]", "B6"),
        ("RC", "C5"),
        ("R2C3:R4C5", "C2:E4"),
        ("R2:R4", "2:4"),
        ("R3", "3:3"),
        ("C3", "C:C"),
        ("C[1]:C[2]", "D:E"),
        ("Rate", "Rate"),
    ],
)
def test_r1c1_references_become_a1(text: str, expected: str) -> None:
    assert to_a1(text, 5, 3) == expected


@pytest.mark.parametrize("text", ["R0C1", "R1048577C1", "R1C16385", "R[-5]C", "RC[-3]", "R2C3:R4"])
def test_r1c1_references_outside_the_sheet_are_ref_errors(text: str) -> None:
    with pytest.raises(FormulaError):
        to_a1(text, 5, 3)


def _book(path: Path, cells: dict[str, str]) -> None:
    workbook = Workbook()
    data = workbook.worksheets[0]
    data.title = "Data"
    data["A1"], data["A2"], data["B2"] = "label", 10, "text"
    data["A3"] = "=A2*2"
    for ref, value in cells.items():
        data[ref] = value
    workbook.create_sheet("My Sheet")["A1"] = 7
    workbook.save(path)


async def test_indirect_reads_cells_named_by_text(call: ToolCall, files: Path) -> None:
    _book(
        files / "calc.xlsx",
        {
            "D1": '=INDIRECT("A2")+1',
            "D2": '=INDIRECT("a"&"3")',
            "D3": "=INDIRECT(\"'My Sheet'!A1\")",
            "D4": '=SUM(INDIRECT("A2:A3"))',
            "D5": '=INDIRECT("R2C1",FALSE)',
            "D6": '=INDIRECT("R[-5]C[-3]",FALSE)',
            "D7": '=INDIRECT("Missing!A1")',
            "D8": '=INDIRECT("nonsense words")',
            "D9": '=ROW(INDIRECT("A5"))+SUMPRODUCT(ROW(INDIRECT("1:3")))',
        },
    )

    data = await read_values(call, "D1:D9")

    assert data["values"] == [[11], [20], [7], [30], [10], ["label"], ["#REF!"], ["#REF!"], [11]]


async def test_indirect_to_another_workbook_is_left_uncalculated(
    call: ToolCall, files: Path
) -> None:
    _book(files / "calc.xlsx", {"D1": '=INDIRECT("[other.xlsx]Sheet1!A1")'})

    data = await read_values(call, "D1")

    assert data["values"] == []
    assert set(data["uncalculated"]) == {"D1"}


async def test_cell_reports_what_excel_would(call: ToolCall, files: Path) -> None:
    _book(
        files / "calc.xlsx",
        {
            "D1": '=CELL("address",A2)',
            "D2": '=CELL("row",B3)*100+CELL("col",B3)',
            "D3": '=CELL("contents",A3)',
            "D4": '=CELL("type",A1)&CELL("type",A2)&CELL("type",Z9)',
            "D5": '=CELL("bogus",A1)',
            "D6": '=CELL("width",A1)',
            "D7": '=CELL("protect",A1)',
            "D8": '=CELL("format",A1)',
        },
    )

    data = await read_values(call, "D1:D8")

    assert data["values"] == [["$A$2"], [302], [20], ["lvb"], ["#VALUE!"], [8], [1], ["G"]]


async def test_cell_filename_names_the_workbook_and_sheet(call: ToolCall, files: Path) -> None:
    _book(files / "calc.xlsx", {"D1": '=CELL("filename",A1)'})

    data = await read_values(call, "D1")

    assert data["values"][0][0].endswith("[calc.xlsx]Data")


async def test_info_answers_only_what_the_file_decides(call: ToolCall, files: Path) -> None:
    _book(
        files / "calc.xlsx",
        {"D1": '=INFO("recalc")', "D2": '=INFO("bogus")', "D3": '=INFO("osversion")'},
    )

    data = await read_values(call, "D1:D3")

    assert data["values"] == [["Automatic"], ["#VALUE!"]]
    assert set(data["uncalculated"]) == {"D3"}
