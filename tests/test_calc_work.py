"""The calculator's work limit counts the work a formula really does: arrays as they are built
and passed on, LAMBDA calls, defined names, and the functions whose work grows faster than
their arguments. Each test fails when that work goes uncounted or takes superlinear time."""

import contextlib
import time
from collections.abc import Callable, Coroutine
from pathlib import Path
from typing import Any

import pytest
from openpyxl import Workbook
from openpyxl.workbook.defined_name import DefinedName

from excel_mcp.calc import engine as engine_module
from excel_mcp.calc.dynamic import is_dynamic_array
from excel_mcp.calc.engine import Engine
from excel_mcp.calc.values import ExcelError, Scalar, UncalculableError
from excel_mcp.xlfn import add_prefixes
from tests.conftest import ToolCall

pytestmark = pytest.mark.anyio

ErrorCall = Callable[..., Coroutine[Any, Any, str]]

# One call does about a hundred million additions while charging a few hundred units of work
# without the fix: the shape of the reported formula, scaled down to keep the tests fast.
AMPLIFIED = "=REDUCE(0,SEQUENCE(100),LAMBDA(a,x,a+SUM(SEQUENCE(10000))))"


def evaluate(formula: str, cells: dict[str, Any] | None = None) -> tuple[Scalar, int]:
    """The formula's value in A1 of a new workbook and the work charged for it."""
    workbook = Workbook()
    sheet = workbook.worksheets[0]
    for ref, value in (cells or {}).items():
        sheet[ref] = value
    sheet["A1"] = add_prefixes(formula)
    engine = Engine(Workbook(), workbook)
    return engine.calculate(sheet, 1, 1), engine.work


def _book(files: Path, cells: dict[str, Any]) -> None:
    workbook = Workbook()
    sheet = workbook.worksheets[0]
    sheet.title = "Data"
    for ref, value in cells.items():
        sheet[ref] = add_prefixes(value) if isinstance(value, str) else value
    workbook.save(files / "book.xlsx")


@pytest.mark.parametrize(
    ("formula", "least"),
    [
        ("=SUM(SEQUENCE(1000))", 1000),
        ("=LET(x,SEQUENCE(1000),SUM(x)+SUM(x)+SUM(x))", 4000),
        ("=SUM(MAP(SEQUENCE(1000),LAMBDA(x,x*2+1)))", 6000),
        ("=SUM(SEQUENCE(1000)*2+1)", 3000),
        ("=SUM(MAP(SEQUENCE(1000),LAMBDA(x,SUM({1,2,3,4,5,6,7,8,9,10}))))", 10_000),
        ("=REDUCE(0,SEQUENCE(10),LAMBDA(a,x,a+SUM(SEQUENCE(1000))))", 10_000),
        ("=SUM(MMULT(SEQUENCE(64,64),SEQUENCE(64,64)))", 64**3 // 32),
        ("=SUM(MINVERSE(SEQUENCE(64,64)+MUNIT(64)))", 64**3 // 32),
        ("=SUM(LINEST(SEQUENCE(100),SEQUENCE(100,20)^0.5))", 100 * 21**3 // 32),
        ('=ROWS(TEXTSPLIT(REPT("ab",1000),SEQUENCE(100)))', 2000 * 100 // 32),
        ("=PRICE(40000,70000,0.05,0.06,100,4)", 300),
        ("=BETA.INV(0.3,1E16,1E16)", 100_000),
        ("=BINOM.DIST(500,1030,0.5,TRUE)", 1000),
        ("=NETWORKDAYS(40000,140000)", 10_000),
        ("=BESSELJ(500,200)", 10_000),
    ],
)
def test_the_work_a_formula_does_is_counted(formula: str, least: int) -> None:
    _, work = evaluate(formula)

    assert work >= least


def test_an_array_built_in_a_loop_is_counted_each_time(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(engine_module, "MAX_WORK", 200_000)

    with pytest.raises(UncalculableError, match="too much to calculate"):
        evaluate(AMPLIFIED)


async def test_read_range_stops_an_amplified_formula(
    call: ToolCall, files: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    _book(files, {"A1": AMPLIFIED})
    monkeypatch.setattr(engine_module, "MAX_WORK", 200_000)

    data = await call("read_range", path="book.xlsx", sheet="Data", range="A1")

    assert data["uncalculated"] == {"A1": "too much to calculate"}


async def test_trace_formula_stops_an_amplified_formula(
    call_error: ErrorCall, files: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    _book(files, {"A1": AMPLIFIED})
    monkeypatch.setattr(engine_module, "MAX_WORK", 200_000)

    message = await call_error(
        "trace_formula", path="book.xlsx", sheet="Data", range="A1", direction="evaluate"
    )

    assert "too much to calculate" in message


async def test_goal_seek_stops_an_amplified_formula(
    call_error: ErrorCall, files: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    _book(files, {"A1": 1, "B1": "=A1" + AMPLIFIED.replace("=", "+", 1)})
    before = (files / "book.xlsx").read_bytes()
    monkeypatch.setattr(engine_module, "MAX_WORK", 200_000)

    message = await call_error(
        "goal_seek", path="book.xlsx", sheet="Data", cell="B1", target=5, changing_cell="A1"
    )

    assert "too much to calculate" in message
    assert (files / "book.xlsx").read_bytes() == before


def test_defined_names_built_from_names_count_every_use() -> None:
    workbook = Workbook()
    sheet = workbook.worksheets[0]
    workbook.defined_names["level0"] = DefinedName("level0", attr_text="1")
    for level in range(1, 8):
        workbook.defined_names[f"level{level}"] = DefinedName(
            f"level{level}", attr_text="+".join([f"level{level - 1}"] * 6)
        )
    sheet["A1"] = "=level7"
    engine = Engine(Workbook(), workbook)

    started = time.perf_counter()
    try:
        value = engine.calculate(sheet, 1, 1)
    except UncalculableError:
        value = None

    assert value in (6.0**7, None)
    assert engine.work >= 6**6
    assert time.perf_counter() - started < 10


def test_the_dynamic_array_check_reads_each_name_once() -> None:
    names = {f"level{level}": "+".join([f"level{level - 1}"] * 40) for level in range(1, 4)}
    texts = {"level0": "A1:A2"} | names
    resolved: list[str] = []

    def resolve(node: Any) -> str | None:
        resolved.append(node.name)
        return texts.get(node.name)

    assert is_dynamic_array("=level3", resolve)
    assert len(resolved) <= 4 * len(texts)


def test_a_formula_scan_counts_the_cells_it_walks() -> None:
    # The IF never reads the range, so only the scan for dependencies walks it.
    count = 300
    cells: dict[str, Any] = {
        f"B{row}": f"=IF(FALSE,SUM($C$1:$C${count}),1)" for row in range(1, count + 1)
    }
    cells |= {f"C{row}": row for row in range(1, count + 1)}
    _, work = evaluate(f"=SUM(B1:B{count})", cells)

    assert work >= count * count


def test_database_functions_count_records_times_criteria() -> None:
    cells: dict[str, Any] = {"C1": "Key", "E1": "Key"}
    cells |= {f"C{row}": row for row in range(2, 1002)}
    cells |= {f"E{row}": row for row in range(2, 102)}
    value, work = evaluate('=DCOUNT(C1:C1001,"Key",E1:E101)', cells)

    assert value == 100
    assert work >= 1000 * 100


@pytest.mark.parametrize(
    "formula",
    [
        '=SEARCH(REPT("a",8000)&"b",REPT("a",16000))',
        '=COUNTIF(B1,"*"&REPT("a",8000)&"b*")',
        '=VALUE(REPT("1",30000)&"x")',
        '=VALUE(REPT("jan",10000))',
        '=IMABS(REPT("1",400)&"x")',
        "=COMBINA(2^30,2^30)",
        "=LCM(SEQUENCE(20000,1,2^52,-1))",
        "=AMORDEGRC(1000,40000,40100,10,1E15,0.15)",
        "=BESSELJ(1,2^31-1)",
        "=BESSELY(5,100000)",
        "=HYPGEOM.DIST(5000,5000,5000,10000,TRUE)",
        '=TEXTJOIN(REPT("x",32767),TRUE,SEQUENCE(3000))',
        '=ARRAYTOTEXT(IF(SEQUENCE(3000),REPT("x",30000)))',
        "=SUM(CHOOSECOLS(SEQUENCE(5000),1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1))",
    ],
)
def test_hostile_arguments_end_quickly(formula: str) -> None:
    started = time.perf_counter()
    with contextlib.suppress(UncalculableError):
        evaluate(formula, {"B1": "a" * 16000})

    assert time.perf_counter() - started < 2


@pytest.mark.parametrize(
    ("formula", "expected"),
    [
        ("=COMBINA(2^30,2^30)", "#NUM!"),
        ("=COMBINA(10,3)", 220.0),
        ("=LCM(4,6,0,2^60)", "#NUM!"),
        ("=LCM(4,6,0)", 0.0),
        ("=AMORDEGRC(1000,40000,40100,10,1E15,0.15)", 0.0),
        ("=BESSELJ(1,2^31-1)", 0.0),
        ("=BESSELY(5,100000)", "#NUM!"),
        ('=TEXTJOIN(REPT("x",32767),TRUE,SEQUENCE(3000))', "#VALUE!"),
        ('=LEN(TEXTJOIN(",",TRUE,SEQUENCE(3)))', 5.0),
        ('=IMABS("3+4i")', 5.0),
        ('=IMABS("1.5.5i")', "#NUM!"),
        ("=GAMMA.INV(0.5,1E17,1)", 1e17),
    ],
)
def test_the_shortcuts_give_excels_results(formula: str, expected: Scalar) -> None:
    value, _ = evaluate(formula)

    assert (value.code if isinstance(value, ExcelError) else value) == expected


def test_a_lambda_calling_itself_is_not_calculated() -> None:
    with pytest.raises(UncalculableError, match="nested too deeply"):
        evaluate("=LET(f,LAMBDA(x,f(x)+1),f(1))")


def test_choosing_columns_is_capped_before_the_array_is_built() -> None:
    picks = ",".join(["1"] * 30)
    with pytest.raises(UncalculableError, match="array too large"):
        evaluate(f"=SUM(CHOOSECOLS(SEQUENCE(5000),{picks}))")


def test_subtotal_leaves_the_rows_of_the_sheet_as_they_are() -> None:
    workbook = Workbook()
    sheet = workbook.worksheets[0]
    sheet["A1"] = "=SUBTOTAL(9,B1:B5000)"
    sheet["B1"] = 1
    Engine(Workbook(), workbook).calculate(sheet, 1, 1)

    assert len(sheet.row_dimensions) < 10
