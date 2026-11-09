"""Calculator behaviour pinned from what Microsoft Excel 16 returned (see formula_golden.json)."""

from collections.abc import Mapping
from typing import Any

from openpyxl import Workbook
from openpyxl.worksheet.filters import FilterColumn, Filters
from openpyxl.worksheet.worksheet import Worksheet

from excel_mcp.calc.engine import Engine
from excel_mcp.calc.values import ExcelError
from excel_mcp.xlfn import add_prefixes


def calculate(
    formula: str,
    cells: Mapping[str, Any] | None = None,
    hidden: tuple[int, ...] = (),
    filtered: bool = False,
) -> Any:
    workbook = Workbook()
    sheet: Worksheet = workbook.active  # pyright: ignore[reportAssignmentType]
    for ref, value in (cells or {}).items():
        sheet[ref] = value
    for row in hidden:
        sheet.row_dimensions[row].hidden = True
    if filtered:
        sheet.auto_filter.ref = "A1:A8"
        sheet.auto_filter.filterColumn.append(FilterColumn(colId=0, filters=Filters(filter=["1"])))
    sheet["Z100"] = add_prefixes(formula)
    value = Engine(Workbook(), workbook).calculate(sheet, 100, 26)
    return value.code if isinstance(value, ExcelError) else value


def test_literals_keep_their_first_15_digits() -> None:
    assert calculate("=GCD(9007199254740993,3)") == 3
    assert calculate('=GCD(VALUE("9007199254740993"),3)') == 3
    assert calculate("=0.99999999999999999*1E15") == 999999999999999
    assert calculate("=1234567890123456.7*10") == 1.23456789012345e16


def test_a_cell_keeps_every_digit_it_holds() -> None:
    assert calculate("=A1-0.123456789012346", {"A1": 0.12345678901234568}) != 0


def test_multiplication_flushes_denormals_and_addition_keeps_them() -> None:
    assert calculate("=2.5E-308*0.5") == 0
    assert calculate("=3E-308-2.5E-308") == 3e-308 - 2.5e-308
    assert calculate("=-0") == 0 and str(calculate("=-0")) == "0.0"
    assert calculate("=EXP(-720)") > 0 and calculate("=EXP(-720)*1") == 0


def test_text_of_a_number_rounds_to_15_digits() -> None:
    assert calculate('="x"&((1000000000000000+234567890123456)*100)') == "x123456789012346000"


def test_approximate_search_runs_excels_bisection_on_unsorted_data() -> None:
    assert calculate("=MATCH(4,{1,2,4,1,7,1,4,1},1)") == 7
    assert calculate("=MATCH(5,{6,4,8,1,8},-1)") == 1
    assert calculate("=XMATCH(4,{5,4,3,4,2,5,9},0,2)") == 2
    assert calculate('=MATCH(3,{"e",3,"c","e",1},1)') == 5
    assert calculate("=VLOOKUP(9,{3;7;1;2;9;2;6},1,TRUE)") == 6


def test_a_blank_key_and_whole_column_sizes() -> None:
    assert calculate("=MATCH(B1,A1:A3,0)", {"A1": 0, "A3": 5}) == 1
    assert calculate("=MATCH(B1,A1:A3,0)", {"A1": 4, "A3": 5}) == "#N/A"
    assert calculate("=ROWS(A:A)") == 1048576
    assert calculate("=COLUMNS(1:1)") == 16384
    assert calculate("=INDEX(A:A,2000)", {"A1": 1}) == 0


def test_undefined_names_and_references() -> None:
    assert calculate("=ROWS(Orders)") == "#NAME?"
    assert calculate("=ISREF(A1:B2)") is True
    assert calculate("=ISREF(A1:B2*2)") is False
    assert calculate("=ISREF(INDEX(A1:C3,2,2))") is True
    assert calculate("=SUM(A1:A3 A2:B2)", {"A1": 1, "A2": 2, "A3": 4}) == 2
    assert (
        calculate("=ISREF(A1:A2 B1:B2)") == "#NULL!" or calculate("=SUM(A1:A2 B1:B2)") == "#NULL!"
    )


def test_subtotal_counts_manually_hidden_rows_but_not_filtered_ones() -> None:
    cells = {f"A{row}": row * 10 for row in range(1, 6)}
    assert calculate("=SUBTOTAL(9,A1:A5)", cells, hidden=(3,)) == 150
    assert calculate("=SUBTOTAL(109,A1:A5)", cells, hidden=(3,)) == 120
    assert calculate("=SUBTOTAL(9,A1:A5)", cells, hidden=(3,), filtered=True) == 120


def test_datedif_month_and_year_cycles() -> None:
    assert calculate('=DATEDIF(DATE(2024,1,31),DATE(2024,3,1),"MD")') == -1
    assert calculate('=DATEDIF(DATE(2023,12,25),DATE(2024,1,5),"MD")') == 11
    assert calculate('=DATEDIF(DATE(2020,3,1),DATE(2024,2,29),"YD")') == 365
    assert calculate('=DATEDIF(DATE(2024,1,1),DATE(2024,3,1),"X")') == "#NUM!"


def test_the_1900_calendar() -> None:
    assert [calculate(f"={f}(0)") for f in ("YEAR", "MONTH", "DAY", "WEEKDAY")] == [1900, 1, 0, 7]
    assert calculate("=DAY(60)") == 29 and calculate("=WEEKDAY(1)") == 1
    assert calculate("=YEAR(2958466)") == "#NUM!"


def test_text_formats() -> None:
    assert calculate('=TEXT(1234.5678,"0.0E+00")') == "1.2E+03"
    assert calculate('=TEXT(1234.5678,"# ?/?")') == "1234 4/7"
    assert calculate('=TEXT(0.5,"##0.0E+0")') == "500.0E-3"
    assert calculate('=TEXT(1234.5,"[Red]0")') == "1235"
    assert calculate('=TEXT(1234.5678,"#,##0.00")') == "1,234.57"
    assert calculate('=TEXT(1.5,"[h]:mm")') == "36:00"
    assert calculate('=TEXT(-1,"[<0]""neg""0;0")') == "neg1"
    assert calculate('=TEXT(1/3,"General")') == "0.333333333"
    assert calculate('=TEXT(0,"yyyy-mm-dd")') == "1900-01-00"


def test_new_functions() -> None:
    assert calculate("=BASE(255,16)") == "FF" and calculate('=DECIMAL("ff",16)') == 255
    assert calculate("=ROMAN(499,4)") == "ID" and calculate('=ARABIC("MCMXCIV")') == 1994
    assert calculate("=DEC2BIN(-1)") == "1111111111" and calculate('=HEX2DEC("FFFFFFFFFF")') == -1
    assert calculate("=BITXOR(5,3)") == 6 and calculate("=BITLSHIFT(1,4)") == 16
    assert calculate("=FACTDOUBLE(7)") == 105 and calculate("=MULTINOMIAL(2,3,4)") == 1260
    assert calculate("=ACOSH(1)") == 0 and calculate("=COT(1)") > 0.64
    assert calculate('=TEXTBEFORE("a,b,c",",",-1)') == "a,b"
    assert calculate('=TEXTAFTER("a,b,c",",",2)') == "c"
    assert calculate('=INDEX(TEXTSPLIT("a,b;c,d",",",";"),2,2)') == "d"
    assert calculate("=INDEX(TOCOL({1,2;3,4}),3)") == 3
    assert calculate("=MDETERM({1,2;3,4})") == -2
    assert calculate("=SUM(TRANSPOSE({1,2,3}))") == 6
    assert calculate("=ADDRESS(5,28,2)") == "AB$5"
    assert calculate("=NORMDIST(0,0,1,TRUE)") == 0.5
