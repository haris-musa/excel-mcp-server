"""A shared formula is translated for each cell that uses it, so the text it expands to is
limited, as the cells are, and counted before any of it is translated."""

import io
from pathlib import Path

import pytest
from openpyxl.formula.translate import Translator

from excel_mcp import lazy_edit, shared_text
from excel_mcp.text_budget import TextBudget
from tests.test_cell_budget import MAIN, call, write_workbook

pytestmark = pytest.mark.anyio

CELLS = 100  # allows 12,800 characters of formula text
MASTER = "+".join(["A1"] * 400)  # 1,199 characters


def sheet_xml(users: int, master: str = '<f{} ref="C1:C1" si="0">') -> str:
    """Rows of a master and ``users`` cells that name its group, whatever its ref says."""
    head = master.format(' t="shared"')
    rows = [f'<row r="1"><c r="A1"><v>1</v></c><c r="C1">{head}{MASTER}</f><v>0</v></c></row>']
    rows += [
        f'<row r="{row}"><c r="A{row}"><v>1</v></c>'
        f'<c r="C{row}"><f t="shared" si="0"/><v>0</v></c></row>'
        for row in range(2, users + 2)
    ]
    return f'<worksheet xmlns="{MAIN}"><sheetData>{"".join(rows)}</sheetData></worksheet>'


@pytest.fixture(autouse=True)
def never_translate(monkeypatch: pytest.MonkeyPatch) -> None:
    def refuse(*_: object, **__: object) -> str:
        raise AssertionError("a formula was translated")

    monkeypatch.setattr(Translator, "translate_formula", refuse)


READS = [
    ("describe_sheet", {}),
    ("write_range", {"at": "A1", "rows": [[1]]}),
    ("read_range", {"range": "A1:C2", "mode": "formulas"}),
    ("find_cells", {"query": "zzz", "mode": "formulas"}),
]


@pytest.mark.parametrize("lazy", [False, True], ids=["whole", "lazy"])
@pytest.mark.parametrize(("tool", "arguments"), READS)
async def test_formulas_that_expand_past_the_limit_are_refused_before_translating(
    files: Path,
    monkeypatch: pytest.MonkeyPatch,
    lazy: bool,
    tool: str,
    arguments: dict[str, object],
) -> None:
    write_workbook(files / "big.xlsx", sheet_xml(30))  # 31 x 1,199 characters
    before = (files / "big.xlsx").read_bytes()
    monkeypatch.setattr(lazy_edit, "LAZY_SHEET_BYTES", 0 if lazy else 1 << 30)
    failed, message = await call(files, tool, CELLS, path="big.xlsx", sheet="Data", **arguments)
    assert failed
    assert "more than 12,800 characters of formula text" in message
    assert "--max-loaded-cells" in message
    assert (files / "big.xlsx").read_bytes() == before


async def test_a_ref_that_understates_the_group_does_not_hide_its_cells(files: Path) -> None:
    write_workbook(files / "big.xlsx", sheet_xml(30, '<f{} si="0" ref="C1">'))
    failed, message = await call(
        files, "read_range", CELLS, path="big.xlsx", sheet="Data", range="A1", mode="formulas"
    )
    assert failed
    assert "characters of formula text" in message


async def test_stored_results_do_not_translate_formulas_and_are_not_counted(files: Path) -> None:
    write_workbook(files / "big.xlsx", sheet_xml(30))
    failed, message = await call(
        files, "find_cells", CELLS, path="big.xlsx", sheet="Data", query="zzz", mode="values"
    )
    assert not failed, message


async def test_a_workbook_within_the_limit_is_opened(
    files: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    monkeypatch.undo()  # translating is what this test is about
    write_workbook(files / "big.xlsx", sheet_xml(8))  # 9 x 1,199 characters
    for tool, arguments in READS:
        failed, message = await call(files, tool, CELLS, path="big.xlsx", sheet="Data", **arguments)
        assert not failed, (tool, message)


@pytest.mark.parametrize("chunk", [7, 64, 1 << 20])
def test_the_text_is_counted_the_same_however_the_sheet_is_read(
    monkeypatch: pytest.MonkeyPatch, chunk: int
) -> None:
    monkeypatch.setattr(shared_text, "_CHUNK", chunk)
    spent = TextBudget(1, "{limit}")
    # A master with an escaped character, a cell that gives other text for the group, and two
    # users: four cells use the first master's text.
    xml = (
        '<c r="A1"><f t="shared" ref="A1:A9" si="3">B1&amp;C1</f></c>'
        '<c r="A2"><f t="shared" si="3">X1+X1+X1</f></c>'
        '<c r="A3"><f t="shared" si="3"/></c><c r="A4"><f si="3" t="shared"/></c>'
        '<c r="A5"><f>NOW()</f></c>'
    )
    shared_text.spend_sheet(io.BytesIO(xml.encode()), spent)
    assert spent.limit - spent.left == 4 * 5


async def test_an_array_formula_that_claims_a_huge_range_is_refused_not_filled(
    files: Path,
) -> None:
    row = (
        '<row r="1"><c r="B1"><f t="array" ref="B1:XFD1048576">A1:A3*2</f><v>0</v></c></row>'
        '<row r="2"><c r="A2"><v>1</v></c></row>'
    )
    write_workbook(
        files / "array.xlsx", f'<worksheet xmlns="{MAIN}"><sheetData>{row}</sheetData></worksheet>'
    )
    failed, message = await call(
        files, "write_range", CELLS, path="array.xlsx", sheet="Data", at="A9", rows=[[1]]
    )
    assert failed
    assert "covers 17,178,820,608 cells" in message
