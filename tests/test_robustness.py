"""Limits and failures found by the robustness sweep: each test fails without its fix."""

from collections.abc import Callable, Coroutine
from pathlib import Path
from typing import Any

import pytest
from openpyxl import Workbook, load_workbook

from excel_mcp.calc.textformat import MAX_FORMAT_LENGTH, format_fixed, format_number
from excel_mcp.calc.values import UncalculableError
from excel_mcp.operations import goal_seek
from tests.conftest import ToolCall

pytestmark = pytest.mark.anyio

ErrorCall = Callable[..., Coroutine[Any, Any, str]]


def _workbook(files: Path, cells: dict[str, Any], name: str = "book.xlsx") -> Path:
    workbook = Workbook()
    sheet = workbook.worksheets[0]
    sheet.title = "Data"
    for ref, value in cells.items():
        sheet[ref] = value
    workbook.save(files / name)
    return files / name


async def test_goal_seek_stops_when_all_iterations_together_take_too_long(
    call_error: ErrorCall, files: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    path = _workbook(files, {"A1": 1, "B1": "=A1^2"})
    before = path.read_bytes()
    monkeypatch.setattr(goal_seek, "MAX_TOTAL_WORK", 20)

    message = await call_error(
        "goal_seek", path="book.xlsx", sheet="Data", cell="B1", target=-1, changing_cell="A1"
    )

    assert "Goal Seek stopped" in message and "20 cell calculations" in message
    assert path.read_bytes() == before


@pytest.mark.parametrize("text_format", ["0." + "0" * 300, "0" * 40_000], ids=["long", "huge"])
def test_a_number_format_beyond_excels_limit_is_not_calculated(text_format: str) -> None:
    with pytest.raises(UncalculableError, match=str(MAX_FORMAT_LENGTH)):
        format_number(1.5, text_format)


def test_digits_beyond_the_decimal_context_are_not_calculated() -> None:
    with pytest.raises(UncalculableError, match="too many digits"):
        format_number(1e300, "0." + "0" * 200)
    with pytest.raises(UncalculableError, match="too many digits"):
        format_fixed(1e20, -1_000_000, grouping=True)


async def test_a_pivot_over_a_hostile_number_format_is_made(call: ToolCall, files: Path) -> None:
    workbook = Workbook()
    sheet = workbook.worksheets[0]
    sheet.title = "Data"
    sheet.append(["Code", "Product", "Units"])
    for index in range(1, 9):
        sheet.append([index * 1e280, f"P{index % 2}", index])
        sheet.cell(index + 1, 1).number_format = "0." + "0" * 250
    workbook.create_sheet("Report")
    workbook.save(files / "book.xlsx")

    result = await call(
        "create_pivot_table",
        path="book.xlsx",
        sheet="Report",
        at="A1",
        source="Data!A1:C9",
        row_fields=["Code", "Product"],
        value_fields=[{"field": "Units", "function": "sum"}],
    )

    assert result["sheet"] == "Report"


async def test_a_quoted_field_may_hold_more_delimiters_than_cells(
    call: ToolCall, call_error: ErrorCall, sample: Path, files: Path
) -> None:
    quoted = '"' + "," * 30_000 + '"'
    nl = chr(10)
    (files / "in.csv").write_text(
        ",".join([quoted] * 4) + ",x" + nl + '"a' + nl + "," * 30_000 + '"' + nl
    )
    result = await call("import_csv", path="sales.xlsx", csv_path="in.csv", sheet="N")
    assert result["cells_written"] == 6
    assert load_workbook(files / "sales.xlsx")["N"]["E1"].value == "x"

    (files / "in.csv").write_text("a" + "," * 100_001 + '"in quotes,,,"\n')
    assert "more than 100,000 cells" in await call_error(
        "import_csv", path="sales.xlsx", csv_path="in.csv", sheet="M"
    )


async def test_evaluate_formula_steps_through_a_long_run_of_operators(
    call: ToolCall, files: Path
) -> None:
    terms = 1_500
    _workbook(files, {"A1": 2, "B1": "=" + "+".join(["A1"] * terms)})

    traced = await call(
        "trace_formula", path="book.xlsx", sheet="Data", range="B1", direction="evaluate"
    )

    assert traced["steps"][-1] == {"expr": "+".join(["A1"] * terms), "value": str(2 * terms)}
    assert traced["more"] > 0


async def test_evaluate_formula_reports_what_it_cannot_calculate_inside_a_run(
    call_error: ErrorCall, files: Path
) -> None:
    _workbook(files, {"B1": '=1+2+CUBEMEMBER("a","b")+3'})

    message = await call_error(
        "trace_formula", path="book.xlsx", sheet="Data", range="B1", direction="evaluate"
    )

    assert "CUBEMEMBER" in message


async def test_describe_sheet_cuts_long_lists_and_notes(call: ToolCall, files: Path) -> None:
    workbook = Workbook()
    sheet = workbook.worksheets[0]
    sheet.title = "Data"
    for row in range(1, 251):
        sheet[f"A{row}"].hyperlink = f"https://example.com/{row}"
        sheet[f"B{row}"].comment = _note("n" * 5_000)
        sheet.merge_cells(start_row=row, start_column=4, end_row=row, end_column=5)
    workbook.save(files / "book.xlsx")

    details = await call("describe_sheet", path="book.xlsx", sheet="Data")

    assert details["more"] == {"merged_ranges": 50, "notes": 50, "hyperlinks": 50}
    assert len(details["notes"]) == len(details["hyperlinks"]) == 200
    text = details["notes"][0]["text"]
    assert len(text) < 1_100 and text.endswith("(4,000 more characters)")


def _note(text: str) -> Any:
    from openpyxl.comments import Comment

    return Comment(text, "me")


async def test_describe_sheet_without_long_lists_has_no_more(call: ToolCall, sample: Path) -> None:
    assert "more" not in await call("describe_sheet", path="sales.xlsx", sheet="Data")


def _damaged(files: Path, source: str, kind: str) -> None:
    import shutil

    from tests.fuzz_support import FIXTURES, damaged_copies

    origin = Path(shutil.copy(FIXTURES / source, files / "origin.xlsx"))
    shutil.copy(damaged_copies(origin, files)[kind], files / "book.xlsx")


@pytest.mark.parametrize("kind", ["bad_sheet_xml", "huge_row", "garbage_sheet"])
@pytest.mark.parametrize(
    ("tool", "arguments"),
    [
        ("describe_workbook", {}),
        ("find_cells", {"query": "a"}),
        ("export_csv", {"sheet": "Sheet1", "csv_path": "out.csv"}),
        ("read_range", {"sheet": "Sheet1"}),
    ],
)
async def test_damaged_sheet_xml_is_a_clear_error(
    call_error: ErrorCall, files: Path, kind: str, tool: str, arguments: dict[str, Any]
) -> None:
    _damaged(files, "excel_pivot_source.xlsx", kind)
    message = await call_error(tool, path="book.xlsx", **arguments)
    assert "Unexpected error" not in message


async def test_a_workbook_without_worksheets_is_a_clear_error(
    call_error: ErrorCall, files: Path
) -> None:
    _damaged(files, "excel_pivot_source.xlsx", "no_sheet")
    message = await call_error(
        "goal_seek", path="book.xlsx", sheet="Data", cell="B1", target=1, changing_cell="A1"
    )
    assert "Unexpected error" not in message


async def test_a_cell_style_that_does_not_exist_is_a_clear_error(
    call_error: ErrorCall, files: Path
) -> None:
    _damaged(files, "excel_pivot_source.xlsx", "bad_styles")
    message = await call_error(
        "goal_seek", path="book.xlsx", sheet="Data", cell="B1", target=1, changing_cell="A1"
    )
    assert "Unexpected error" not in message


@pytest.mark.parametrize("axis", ["rows", "columns"])
@pytest.mark.parametrize("tool", ["insert_rows_or_columns", "delete_rows_or_columns"])
async def test_lines_past_the_last_row_or_column_are_refused_before_saving(
    call_error: ErrorCall, files: Path, axis: str, tool: str
) -> None:
    path = _workbook(files, {"A1": 1})
    before = path.read_bytes()
    limit = 1_048_576 if axis == "rows" else 16_384

    message = await call_error(
        tool, path="book.xlsx", sheet="Data", axis=axis, start=limit, count=2
    )

    assert "past Excel's last" in message
    assert path.read_bytes() == before


async def test_inserting_rows_next_to_an_unformatted_cell(call: ToolCall, files: Path) -> None:
    import shutil

    from tests.fuzz_support import FIXTURES

    shutil.copy(FIXTURES / "excel_controls.xlsx", files / "book.xlsx")

    result = await call(
        "insert_rows_or_columns", path="book.xlsx", sheet="Data", axis="rows", start=2, count=3
    )

    assert result["range"] == "2:4"
