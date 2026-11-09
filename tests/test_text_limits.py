# pyright: reportOptionalMemberAccess=false, reportAttributeAccessIssue=false
"""Work whose size comes from what cells hold, not from how many there are, is bounded."""

from pathlib import Path
from typing import Any, cast

import pytest
from mcp import Client
from openpyxl import Workbook, load_workbook
from openpyxl.worksheet.worksheet import Worksheet

from excel_mcp.calc import engine
from excel_mcp.config import Limits, Settings
from excel_mcp.operations import replace, text_split
from excel_mcp.server import create_server
from tests.conftest import error_text

pytestmark = pytest.mark.anyio

SMALL = Limits(max_cells=10, max_read_cells=10)  # 2,560 characters of text per call


def make(files: Path, name: str, rows: list[list[Any]]) -> Path:
    workbook = Workbook()
    for row in rows:
        workbook.active.append(row)  # pyright: ignore[reportOptionalMemberAccess]
    path = files / name
    workbook.save(path)
    return path


async def call(files: Path, limits: Limits, tool: str, **arguments: Any) -> tuple[bool, str, Any]:
    async with Client(create_server(Settings(allowed_dirs=[files], limits=limits))) as client:
        result = await client.call_tool(tool, arguments)
    return result.is_error, error_text(result), result.structured_content


async def refused(files: Path, limits: Limits, tool: str, **arguments: Any) -> str:
    is_error, text, _ = await call(files, limits, tool, **arguments)
    assert is_error, text
    return text


def split(**transform: Any) -> dict[str, Any]:
    return {"operation": "text_to_columns", **transform}


async def test_text_to_columns_stops_at_the_limit_instead_of_splitting_everything(
    files: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    make(files, "a.xlsx", [["," * 2000]] * 3)
    made = 0
    real = text_split._pieces

    def counting(*args: Any) -> Any:
        nonlocal made
        for piece in real(*args):
            made += 1
            yield piece

    monkeypatch.setattr(text_split, "_pieces", counting)
    message = await refused(
        files, SMALL, "transform_range", path="a.xlsx", sheet="Sheet", range="A1:A3",
        transform=split(delimiters=["comma"]),
    )  # fmt: skip
    assert "Splitting would fill more than 10 cells." in message
    assert made == 11


async def test_text_to_columns_counts_the_text_it_reads(files: Path) -> None:
    make(files, "a.xlsx", [["x" * 1000]] * 5)
    message = await refused(
        files, SMALL, "transform_range", path="a.xlsx", sheet="Sheet", range="A1:A5",
        transform=split(delimiters=["comma"]),
    )  # fmt: skip
    assert "more than 2,560 characters of text" in message


@pytest.mark.parametrize("widths", [[0, 3], [-2], [1] * 20_000])
async def test_text_to_columns_widths_must_be_positive_and_few(
    files: Path, widths: list[int]
) -> None:
    make(files, "a.xlsx", [["abcdef"]])
    message = await refused(
        files, Limits(), "transform_range", path="a.xlsx", sheet="Sheet", range="A1",
        transform=split(fixed_widths_chars=widths),
    )  # fmt: skip
    assert "fixed_widths_chars" in message


async def test_text_to_columns_still_splits_normally(files: Path) -> None:
    path = make(files, "a.xlsx", [["a,b,,d"], ["1,2"]])
    is_error, text, _ = await call(
        files, Limits(), "transform_range", path="a.xlsx", sheet="Sheet", range="A1:A2",
        transform=split(delimiters=["comma"], merge_delimiters=True),
    )  # fmt: skip
    assert not is_error, text
    sheet = load_workbook(path).active
    assert [[c.value for c in row] for row in sheet.iter_rows()] == [
        ["a", "b", "d"],
        [1, 2, None],
    ]


async def test_a_read_page_ends_where_its_text_reaches_the_limit(files: Path) -> None:
    make(files, "a.xlsx", [["y" * 1000]] * 6)
    is_error, text, data = await call(files, SMALL, "read_range", path="a.xlsx", sheet="Sheet")
    assert not is_error, text
    assert data["range"] == "A1:A2"
    assert data["next_range"] == "A3:A6"


async def test_a_row_longer_than_a_page_of_text_is_refused(files: Path) -> None:
    make(files, "a.xlsx", [["y" * 1000, "z" * 1000, "w" * 1000]])
    message = await refused(files, SMALL, "read_range", path="a.xlsx", sheet="Sheet")
    assert "Read fewer columns" in message


async def test_export_csv_refuses_more_text_than_the_limit(files: Path) -> None:
    make(files, "a.xlsx", [["y" * 1000]] * 5)
    message = await refused(
        files, SMALL, "export_csv", path="a.xlsx", sheet="Sheet", csv_path="o.csv"
    )
    assert "characters of text" in message
    assert not (files / "o.csv").exists()


async def test_find_cells_returns_a_page_of_text_and_refuses_an_endless_scan(
    files: Path,
) -> None:
    make(files, "a.xlsx", [["match " + "y" * 4000]] * 5)
    _, _, data = await call(files, SMALL, "find_cells", path="a.xlsx", query="match")
    assert data["truncated"] is True
    assert len(data["matches"]["Sheet"]) == 1
    message = await refused(
        files, Limits(max_loaded_cells=10), "find_cells", path="a.xlsx", query="zzz"
    )
    assert "more than 2,560 characters of text" in message


async def test_replace_works_out_a_shared_text_once_and_refuses_growth(
    files: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    path = make(files, "a.xlsx", [["a,b"]] * 40)
    worked = 0
    real = replace._replaced

    def counting(*args: Any, **kwargs: Any) -> str:
        nonlocal worked
        worked += 1
        return real(*args, **kwargs)

    monkeypatch.setattr(replace, "_replaced", counting)
    is_error, text, _ = await call(
        files, Limits(), "replace_cells", path="a.xlsx", query=",", replacement=";"
    )
    assert not is_error, text
    assert worked == 1
    assert [row[0].value for row in load_workbook(path).active.iter_rows()] == ["a;b"] * 40

    message = await refused(
        files, SMALL, "replace_cells", path="a.xlsx", query=";", replacement="x" * 5000
    )
    assert "would add more than 2,560 characters" in message


async def test_replace_refuses_a_result_longer_than_a_cell(files: Path) -> None:
    make(files, "a.xlsx", [["a" * 30_000]])
    message = await refused(
        files, Limits(), "replace_cells", path="a.xlsx", query="a", replacement="bb"
    )
    assert "32,767" in message


async def test_sort_and_remove_duplicates_count_the_text_they_compare(files: Path) -> None:
    make(files, "a.xlsx", [["Header"], *[["k" * 1000 + str(i)] for i in range(5)]])
    sort = await refused(
        files, SMALL, "sort_range", path="a.xlsx", sheet="Sheet", range="A1:A6",
        sort_by=[{"column": "Header"}],
    )  # fmt: skip
    assert "Sorting would compare more than 2,560 characters" in sort
    dedupe = await refused(
        files, SMALL, "transform_range", path="a.xlsx", sheet="Sheet", range="A1:A6",
        transform={"operation": "remove_duplicates"},
    )  # fmt: skip
    assert "Comparing rows would read more than 2,560 characters" in dedupe


async def test_import_csv_refuses_a_record_with_too_many_fields_before_reading_it(
    files: Path,
) -> None:
    make(files, "a.xlsx", [["x"]])
    (files / "wide.csv").write_text("," * 50 + "\n")
    message = await refused(
        files,
        SMALL,
        "import_csv",
        path="a.xlsx",
        csv_path="wide.csv",
        sheet="Sheet",
        delimiter="comma",
    )
    assert "more than 10 cells" in message


async def test_a_pivot_source_counts_its_text(files: Path) -> None:
    make(files, "a.xlsx", [["Key", "Note"], *[[f"k{i}", "n" * 1000] for i in range(5)]])
    message = await refused(
        files,
        Limits(max_pivot_source_cells=12),
        "create_pivot_table",
        path="a.xlsx",
        sheet="Sheet",
        at="E1",
        source="A1:B6",
        row_fields=["Key"],
        value_fields=[{"field": "Note", "function": "count"}],
    )
    assert "more than 3,072 characters of text" in message


async def test_a_pivot_table_is_sized_before_its_cells_are_made(
    files: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    make(files, "a.xlsx", [["A", "B", "C"], *[[f"a{i}", f"b{i}", 1] for i in range(30)]])

    def not_called(*args: Any) -> Any:
        raise AssertionError("the table was built before it was checked")

    monkeypatch.setattr("excel_mcp.operations.pivot.render", not_called)
    message = await refused(
        files,
        Limits(max_cells=200),
        "create_pivot_table",
        path="a.xlsx",
        sheet="Sheet",
        at="F1",
        source="A1:C31",
        row_fields=["A"],
        column_fields=["B"],
        value_fields=[{"field": "C"}],
    )
    assert "at most 200" in message


async def test_a_table_slicer_does_not_make_the_cells_of_a_huge_declared_table(
    files: Path,
) -> None:
    from openpyxl.worksheet.table import Table

    workbook = Workbook()
    sheet = workbook.active
    sheet.append(["Name", "Qty"])
    sheet.append(["n", 1])
    sheet.add_table(Table(displayName="T", ref="A1:B5000"))
    workbook.save(files / "a.xlsx")
    message = await refused(
        files, Limits(max_cells=100), "add_slicer", path="a.xlsx", sheet="Sheet",
        target={"sheet": "Sheet", "name": "T"}, field="Name", at="E2", selected_items=["n"],
    )  # fmt: skip
    assert "at most 100" in message


async def test_a_totals_row_does_not_make_the_cells_of_a_huge_declared_table(
    files: Path,
) -> None:
    from openpyxl.worksheet.table import Table

    workbook = Workbook()
    sheet = workbook.active
    sheet.append(["Name", "Qty"])
    sheet.append(["n", 1])
    sheet.add_table(Table(displayName="T", ref="A1:B50000"))
    workbook.save(files / "a.xlsx")
    is_error, text, _ = await call(
        files, Limits(), "set_table", path="a.xlsx", sheet="Sheet", name="T",
        options={"totals_row": True},
    )  # fmt: skip
    assert not is_error, text
    cells = load_workbook(files / "a.xlsx").active._cells  # pyright: ignore[reportOptionalMemberAccess]
    assert len(cells) < 20


async def test_autofitting_many_columns_reads_the_sheet_once(files: Path) -> None:
    make(files, "a.xlsx", [["y" * 1000, "z" * 3000]] * 2)
    message = await refused(
        files, SMALL, "set_sheet_layout", path="a.xlsx", sheet="Sheet",
        layout={"autofit_columns": ["A", "B", "A", "B"]},
    )  # fmt: skip
    assert "Fitting columns would read more than 2,560 characters" in message
    is_error, text, _ = await call(
        files, Limits(), "set_sheet_layout", path="a.xlsx", sheet="Sheet",
        layout={"autofit_columns": ["A"] * 1000},
    )  # fmt: skip
    assert not is_error, text


async def test_a_weekday_series_with_a_huge_step_is_fast_and_exact(files: Path) -> None:
    import datetime as dt

    path = make(files, "a.xlsx", [[dt.date(2026, 11, 7)], [None], [None]])
    spec = {"operation": "fill", "direction": "down", "series": "date", "unit": "weekday"}
    is_error, text, _ = await call(
        files, Limits(), "transform_range", path="a.xlsx", sheet="Sheet", range="A1:A3",
        transform={**spec, "step": 7},
    )  # fmt: skip
    assert not is_error, text
    values = [row[0].value.date() for row in load_workbook(path).active.iter_rows()]
    assert values == [dt.date(2026, 11, 7), dt.date(2026, 11, 17), dt.date(2026, 11, 26)]
    message = await refused(
        files, Limits(), "transform_range", path="a.xlsx", sheet="Sheet", range="A1:A3",
        transform={**spec, "step": 10**9},
    )  # fmt: skip
    assert "last date" in message


async def test_a_copy_that_multiplies_a_text_past_what_can_be_reopened_is_refused(
    files: Path,
) -> None:
    path = make(files, "a.xlsx", [["w" * 1000]] * 4)
    before = path.read_bytes()
    limits = Limits(max_file_bytes=6000, max_unpack_factor=1)
    message = await refused(
        files, limits, "copy_range", path="a.xlsx", sheet="Sheet", range="A1:A4", at="B1"
    )
    assert "expands to more than 6,000 bytes" in message
    assert path.read_bytes() == before


async def test_the_calculator_counts_the_text_of_the_cells_it_reads(
    files: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    make(
        files,
        "a.xlsx",
        [["y" * 5000] for _ in range(10)] + [['=COUNTIF(A1:A10,"*zzz*")']],
    )
    monkeypatch.setattr(engine, "MAX_WORK", 1000)
    _, _, data = await call(
        files, Limits(), "read_range", path="a.xlsx", sheet="Sheet", range="A11"
    )
    assert data["uncalculated"] == {"A11": "too much to calculate"}


def test_a_line_with_too_many_delimiters_is_refused_before_it_is_read_as_fields() -> None:
    import io

    from excel_mcp.errors import LimitExceededError
    from excel_mcp.operations.csv_import import _lines_within

    lines = _lines_within(io.StringIO("a,b\n" + "," * 10 + "\n"), ",", 10)
    assert next(lines) == "a,b\n"
    with pytest.raises(LimitExceededError, match="more than 10 cells"):
        next(lines)


def test_checking_a_huge_declared_table_makes_no_cells() -> None:
    from openpyxl.worksheet.table import Table

    from excel_mcp.operations.table_options import _is_numeric

    sheet = cast(Worksheet, Workbook().active)
    sheet.append(["Qty"])
    sheet.append([1])
    table = Table(displayName="T", ref="A1:A500000")
    sheet.add_table(table)
    assert _is_numeric(sheet, table, 1)
    assert len(sheet._cells) == 2
