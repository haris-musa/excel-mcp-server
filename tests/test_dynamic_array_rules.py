"""Which formulas are stored as dynamic array formulas, and the calculator's array evaluation.

The tables list what Microsoft Excel itself saved for each formula entered with Formula2
(a formula that spills or needs array evaluation gets ``t="array"`` and ``cm``; any other
is stored as an ordinary formula). Excel 16.0 build 20430 on 2026-10-08.
"""

import pytest
from openpyxl import Workbook

from excel_mcp.calc.dynamic import is_dynamic_array
from excel_mcp.calc.engine import Engine
from excel_mcp.calc.values import ExcelError

DYNAMIC = [
    "=A1:A5",
    "=A1:B2+1",
    "=-A1:A5",
    "=A1:A5%",
    '=A1:A5&"x"',
    "=A1:A5=3",
    '=A1:A5<>""',
    "=SUM(A1:A5*2)",
    "=SUM(A1:A5*B1:B5)",
    "=SUMPRODUCT(A1:A5*B1:B5)",
    "=SUMPRODUCT(--(A1:A5>1))",
    "=MAX(A1:A5-1)",
    "=SUM(--(A1:A5>1))",
    "=SUM(IF(A1:A5>1,1,0))",
    "=AND(A1:A5>1)",
    "=AGGREGATE(9,6,A1:A5/(A1:A5>1))",
    '=TEXTJOIN(",",1,IF(A1:A5>1,A1:A5,""))',
    "=IF(A1>1,A1:A5,B1:B5)",
    "=IFERROR(A1:A5,0)",
    "=CHOOSE(2,A1:A5,B1:B5)",
    "=CHOOSE({1,2},A1,B1)",
    "=FILTER(A1:A5,A1:A5>3)",
    "=FILTER(A1:A5,A1:A5>100)",
    "=SORT(A1:A5)",
    "=SORT(A1:A5,1,-1)",
    "=SORTBY(A1:A5,B1:B5,-1)",
    "=UNIQUE(A1:A5)",
    "=SEQUENCE(3)",
    "=SEQUENCE(1)",
    "=SEQUENCE(2,2)",
    "=RANDARRAY(2,2)",
    "=TRANSPOSE(A1:B2)",
    "=TRANSPOSE(A1)",
    "=SUM(TRANSPOSE(A1:B2))",
    "=TAKE(A1:B5,2)",
    "=MMULT(A1:B2,A1:B2)",
    "=SORT(A1)",
    "=FILTER(A1,A1>0)",
    "=UNIQUE(A1)",
    "=SORT(UNIQUE(A1:A5))",
    "=SEQUENCE(ROWS(A1:A5))",
    "=XLOOKUP(3,A1:A5,A1:B5)",
    "=XLOOKUP(A1:A3,A1:A5,B1:B5)",
    "=INDEX(A1:B5,0,2)",
    "=INDEX(A1:B5,2,0)",
    "=INDEX(A1:A5,0)",
    "=INDEX(A1:B5,0,1)",
    "=INDEX(A1:B5,MATCH(3,A1:A5,0),0)",
    "=INDEX(A1:A5,{1,2})",
    "=SUM(INDEX(A1:A5,{1,2}))",
    "=OFFSET(A1,0,0,2,1)",
    "=OFFSET(A1,0,0,3)",
    "=OFFSET(A1,0,0,B1,1)",
    "=ROW(A1:A5)",
    "=COLUMN(A1:C1)",
    "=SUM(ROW(A1:A5))",
    "=SUM(ROW(A1:A5)^2)",
    "=SMALL(A1:A5,ROW(A1:A3))",
    "=LARGE(A1:A5,{1,2})",
    "=SUM(LARGE(A1:A5,{1,2}))",
    "=MATCH({1,2},A1:A5,0)",
    "=VLOOKUP(A1:A3,A1:B5,2,0)",
    "=COUNTIF(A1:A5,A1:A5)",
    "=LEN(A1:A5)",
    '=TEXT(A1:A5,"0")',
    "=ISNUMBER(A1:A5)",
    "=ABS(A1:A5)",
    "=SUM(ABS(A1:A5))",
    "=ROUND(SORT(A1:A5),0)",
    "=LEN(SORT(A1:A5))",
    "=IF(TRUE,SORT(A1:A5),0)",
    "=CHOOSE(2,SORT(A1:A5),B1:B5)",
    "=SORT(A1:A5)+1",
    "=SUM(SORT(A1:A5)*2)",
    "=LET(x,A1:A5,x*2)",
    "=LET(a,A1:A5,b,a*2,SUM(b))",
    "=MAP(A1:A5,LAMBDA(x,x*2))",
    "=BYROW(A1:B5,LAMBDA(r,SUM(r)))",
    "=REDUCE(0,A1:A5,LAMBDA(a,b,a+b))",
    "=SCAN(0,A1:A5,LAMBDA(a,b,a+b))",
    "=INDEX(FILTER(A1:B5,A1:A5>2),1,2)",
    '=COUNTA(FILTER(A1:A5,B1:B5>10,""))',
    "=A1:A5+0",
    "=SUM(A1:A5+1)",
]

ORDINARY = [
    "=SUM(A1:A5)",
    "=INDEX(A1:A5,2)",
    "=VLOOKUP(A1,A1:B5,2,0)",
    "=XLOOKUP(3,A1:A5,B1:B5)",
    "=A1*2",
    "=SUBTOTAL(9,A1:A5)",
    '=COUNTIF(A1:A5,">1")',
    "=LOOKUP(2,A1:A5,B1:B5)",
    "=MATCH(3,A1:A5,0)",
    "=SUM(OFFSET(A1,0,0,3,1))",
    "=INDEX(A1:A5,MATCH(3,A1:A5,0))",
    "=SUM(A:A)",
    "=SUM(SORT(A1:A5))",
    "=SUM(SEQUENCE(3))",
    "=COUNTA(UNIQUE(A1:A5))",
    "=MAX(INDEX(A1:B5,0,2))",
    "=SUMPRODUCT(A1:A5)",
    "=SUMPRODUCT(A1:A5,B1:B5)",
    "=SUM(A1:B2)",
    "=ROW(A1)",
    "=OFFSET(A1,0,0,1,1)",
    "=SUM(CHOOSE({1,2},A1,B1))",
    "=XMATCH(3,A1:A5)",
    "=LEN(A1)",
    '=SUMIF(A1:A5,">1",B1:B5)',
    "=AVERAGE(A1:A5)",
    "=COUNTA(A:A)",
    "=LET(x,A1:A5,SUM(x))",
    "=N(A1:A5)",
    "=SUM(N(A1:A5))",
    "=IFERROR(A1,0)",
    "=HLOOKUP(3,A1:B5,2,0)",
    "=SUM(TAKE(A1:B5,2))",
    "=SUM(A1,B1:B5)",
    "=MAX(A1:A5,B1:B5)",
    "=SUM(MMULT(A1:B2,A1:B2))",
    '=TEXTJOIN(",",1,A1:A5)',
    "=CONCAT(A1:A5)",
    "=ROWS(A1:A5)",
    "=COUNT(A1:A5)",
    "=ROWS(SORT(A1:A5))",
    "=MAX(SORT(A1:A5))",
    "=LARGE(SORT(A1:A5),1)",
    "=SUM(FILTER(A1:A5,A1:A5))",
    "=MATCH(3,SORT(A1:A5),0)",
    "=VLOOKUP(3,SORT(A1:B5),2,0)",
    "=SUM(INDEX(A1:B5,0,2))",
    "=SUM(OFFSET(A1,0,0,2,1))",
    "=ISNUMBER(MATCH(3,A1:A5,0))",
    '=COUNTIFS(A1:A5,">1",B1:B5,"<50")',
    "=SUM(A1:A5,SORT(A1:A5))",
    "=A1&B1",
    '=SUM(A1:A5)&"x"',
    "=INDEX(B1:B5,MATCH(3,A1:A5,0))",
    "=INDEX(A1:B5,1,MATCH(20,B1:B5,0))",
    "=INDEX(A1:B5,MATCH(3,A1:A5,0),2)",
    "=OFFSET(A1,A1,0)",
    "=IFERROR(INDEX(B1:B5,MATCH(9,A1:A5,0)),0)",
    "=IF(A1>1,B1,0)",
    "=SUMIFS(B1:B5,A1:A5,A1)",
    "=SUM(INDEX(A1:B5,0,0))",
    "=INDEX(A1:B5,2)",
    "=INDEX(A1:A5,2,1)",
    "=LARGE(A1:A5,A1)",
    '=XLOOKUP(A1,A1:A5,B1:B5,"none")',
    "=COUNTIF(A1:A5,A1)",
    "=INDEX(A:A,1)",
    "=MATCH(A1,A1:A5,0)",
    '=TEXT(A1,"0")',
    "=SUM(A1:A5)/COUNT(A1:A5)",
    '=TEXTJOIN(",",TRUE,SORT(A1:A5))',
    "=ROWS(UNIQUE(A1:A5))",
    "=AVERAGE(INDEX(A1:B5,0,2))",
    "=IF(A1>1,SUM(A1:A5),0)",
    "=SUM(A1:A5)*2",
    "=INDEX(A1:A5,1)+A2",
    "=MAX(A1:A5)-MIN(A1:A5)",
    "=ROUND(A1,0)",
    "=A1+B1",
]


@pytest.mark.parametrize("formula", DYNAMIC)
def test_formulas_excel_saves_as_dynamic_arrays(formula: str) -> None:
    assert is_dynamic_array(formula, lambda name: None)


@pytest.mark.parametrize("formula", ORDINARY)
def test_formulas_excel_saves_as_ordinary_formulas(formula: str) -> None:
    assert not is_dynamic_array(formula, lambda name: None)


def test_names_are_looked_through() -> None:
    names = {"Prices": "Data!$B$2:$B$9", "One": "Data!$B$2"}
    resolve = lambda name: names.get(name.name)  # noqa: E731

    assert is_dynamic_array("=Prices*2", resolve)
    assert not is_dynamic_array("=SUM(Prices)", resolve)
    assert not is_dynamic_array("=One*2", resolve)


def _engine(formula: str) -> tuple[Engine, object]:
    workbook = Workbook()
    sheet = workbook.worksheets[0]
    for row in range(1, 6):
        sheet.cell(row, 1, row)
        sheet.cell(row, 2, row * 10)
    return Engine(Workbook(), workbook), sheet


@pytest.mark.parametrize(
    ("formula", "expected"),
    [
        ("=SUM(IF(A1:A5>2,A1:A5,0))", [[12]]),
        ("=IF(A1:A5>2,B1:B5,-1)", [[-1], [-1], [30], [40], [50]]),
        ('=IF(A1:A5>2,"big")', [["FALSE"], ["FALSE"], ["big"], ["big"], ["big"]]),
        ("=IF(A1>0,A1:A3,0)", [[1], [2], [3]]),
        ("=A1:A3*2", [[2], [4], [6]]),
        ("=_xlfn._xlws.SORT(B1:B5,1,-1)", [[50], [40], [30], [20], [10]]),
        ("=_xlfn.XLOOKUP(3,A1:A5,A1:B5)", [[3, 30]]),
    ],
)
def test_the_calculator_evaluates_dynamic_arrays(
    formula: str, expected: list[list[object]]
) -> None:
    engine, sheet = _engine(formula)

    grid = engine.spill(sheet, 1, 4, formula)  # type: ignore[arg-type]

    shown = [[("FALSE" if item is False else item) for item in row] for row in grid.rows]
    assert shown == expected


def test_a_range_formula_in_a_cell_shows_its_first_value_when_not_array_entered() -> None:
    engine, sheet = _engine("")
    sheet["D3"] = "=A1:A5"  # type: ignore[index]

    assert engine.calculate(sheet, 3, 4) == 3  # type: ignore[arg-type]


def test_errors_inside_an_array_stay_errors() -> None:
    engine, sheet = _engine("")

    grid = engine.spill(sheet, 1, 4, "=1/(A1:A3-2)")  # type: ignore[arg-type]

    assert [type(row[0]) is ExcelError for row in grid.rows] == [False, True, False]
