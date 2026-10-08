import pytest

from excel_mcp.operations.shifting import Shifter
from excel_mcp.package.lines import Axis, LineEdit
from excel_mcp.refs import CellRange


def shifter(axis: Axis, at: int, count: int = 1, *, delete: bool = False) -> Shifter:
    return Shifter(LineEdit("S", axis, at, count, delete), {})


@pytest.mark.parametrize(
    ("formula", "expected"),
    [
        ("=A1+A5", "=A1+A7"),
        ("=SUM(A2:A8)", "=SUM(A2:A10)"),
        ("=SUM(A5:A8)", "=SUM(A7:A10)"),
        ("=SUM(A1:A3)", "=SUM(A1:A3)"),
        ("=SUM(A1:A4)", "=SUM(A1:A6)"),
        ("=$A$5+B$5+$B5", "=$A$7+B$7+$B7"),
        ("=SUM(A:A)+SUM(3:5)", "=SUM(A:A)+SUM(3:7)"),
        ("=SUM(A1:A1048576)", "=SUM(A1:A1048576)"),
        ('=A5&"A5"', '=A7&"A5"'),
        ("=Other!A5+S!A5+'S'!A5", "=Other!A5+S!A7+'S'!A7"),
        ("=SUM(S:T!A5)", "=SUM(S:T!A5)"),
        ("=Tax+Sheet1!Rate+Tbl[Col]", "=Tax+Sheet1!Rate+Tbl[Col]"),
        ("=A3:INDEX(B:B,2)", "=A3:INDEX(B:B,2)"),
        ("=A4:INDEX(B:B,2)", "=A6:INDEX(B:B,2)"),
        ("=SUM(A5:A3)", "=SUM(A3:A7)"),
        ("=SUM(A1:B3 A2:C6)", "=SUM(A1:B3 A2:C8)"),
    ],
)
def test_insert_rows(formula: str, expected: str) -> None:
    assert shifter("rows", 4, 2).formula(formula, "S") == expected


@pytest.mark.parametrize(
    ("formula", "expected"),
    [
        ("=A3", "=#REF!"),
        ("=A2+A5", "=A2+A3"),
        ("=SUM(A3:A4)", "=SUM(#REF!)"),
        ("=SUM(A1:A3)", "=SUM(A1:A2)"),
        ("=SUM(A4:A8)", "=SUM(A3:A6)"),
        ("=SUM(A1:A9)", "=SUM(A1:A7)"),
        ("=SUM(A5:A1048576)", "=SUM(A3:A1048576)"),
        ("=SUM(2:6)", "=SUM(2:4)"),
        ("=SUM(3:4)", "=SUM(#REF!)"),
        ("=Other!A3+S!A3", "=Other!A3+S!#REF!"),
        ("=SUM(B:C)", "=SUM(B:C)"),
    ],
)
def test_delete_rows(formula: str, expected: str) -> None:
    assert shifter("rows", 3, 2, delete=True).formula(formula, "S") == expected


@pytest.mark.parametrize(
    ("formula", "expected"),
    [
        ("=B2+C2", "=#REF!+B2"),
        ("=SUM(B:D)", "=SUM(B:C)"),
        ("=SUM(B:B)", "=SUM(#REF!)"),
        ("=SUM(A1:B2)", "=SUM(A1:A2)"),
        ("=SUM(A:XFD)", "=SUM(A:XFD)"),
        ("=SUM(E:XFD)", "=SUM(D:XFD)"),
        ("=SUM(3:4)", "=SUM(3:4)"),
    ],
)
def test_delete_columns(formula: str, expected: str) -> None:
    assert shifter("columns", 2, 1, delete=True).formula(formula, "S") == expected


def test_references_on_other_sheets_are_left_alone() -> None:
    assert shifter("rows", 1).formula("=A5+S!A5", "Other") == "=A5+S!A6"
    assert shifter("rows", 1).operand("S!$A$1:$B$5", "") == "S!$A$2:$B$6"


def test_references_to_deleted_tables_and_columns() -> None:
    dead = {"orders": frozenset({"qty"}), "gone": None}
    edit = LineEdit("S", "columns", 2, 1, delete=True)
    shifted = Shifter(edit, dead)
    assert shifted.formula("=SUM(Orders[Qty])", "S") == "=SUM(#REF!)"
    assert shifted.formula("=SUM(Orders[[Qty]:[Cost]])", "S") == "=SUM(#REF!)"
    assert shifted.formula("=SUM(Orders[Cost])", "S") == "=SUM(Orders[Cost])"
    assert shifted.formula("=Gone[Any]", "S") == "=#REF!"


def test_spans() -> None:
    insert = LineEdit("S", "rows", 5, 2, delete=False)
    assert insert.span(1, 4) == (1, 4)
    assert insert.span(4, 5) == (4, 7)
    assert insert.span(5, 9) == (7, 11)
    delete = LineEdit("S", "rows", 5, 2, delete=True)
    assert delete.span(5, 6) is None
    assert delete.span(4, 5) == (4, 4)
    assert delete.span(6, 9) == (5, 7)
    assert delete.index(6) is None
    assert delete.index(7) == 5
    assert delete.range(CellRange(1, 1, 8, 3)) == CellRange(1, 1, 6, 3)


def test_cuts() -> None:
    insert = LineEdit("S", "rows", 5, 1, delete=False)
    assert insert.cuts(3, 5)
    assert not insert.cuts(5, 8)
    assert not insert.cuts(1, 4)
    delete = LineEdit("S", "rows", 5, 2, delete=True)
    assert delete.cuts(4, 5)
    assert delete.cuts(6, 9)
    assert not delete.cuts(5, 6)
    assert not delete.cuts(7, 9)
