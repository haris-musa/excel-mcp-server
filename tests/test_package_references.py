"""Preserved content follows rows and columns that are inserted or deleted.

Expected values are what Microsoft Excel itself saved after the same edit of the same
Excel-authored fixture (see tests/fixtures/README.md).
"""

import re
from pathlib import Path

import pytest
from openpyxl import Workbook, load_workbook

from excel_mcp import package
from excel_mcp.package import extensions as ext
from excel_mcp.package import state_of
from excel_mcp.package.lines import LineEdit
from excel_mcp.package.model import SheetPackage
from excel_mcp.package.references import rewrite_formulas, rewrite_lines
from tests.package_support import copy_fixture

_REFERENCE = re.compile(r"(?:(\w+)!)?\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?")


def _formula(edit: LineEdit):
    """A stand-in for the caller's formula rewriting: moves plain references to ``edit.sheet``."""

    def move(text: str, host: str) -> str:
        def one(match: re.Match[str]) -> str:
            sheet = match[1] or host
            if sheet != edit.sheet:
                return match[0]
            return _moved(match, edit)

        return _REFERENCE.sub(one, text)

    return move


def _moved(match: re.Match[str], edit: LineEdit) -> str:
    sheet, first_col, first_row, last_col, last_row = match.groups()
    prefix = f"{sheet}!" if sheet else ""
    cells = [(first_col, int(first_row)), (last_col or first_col, int(last_row or first_row))]
    rows = edit.axis == "rows"
    lines = [row if rows else _number(col) for col, row in cells]
    span = edit.span(lines[0], lines[-1])
    if span is None:
        return f"{prefix}#REF!"
    text = []
    for (col, row), line in zip(cells, span, strict=True):
        text.append(f"{col}{line}" if rows else f"{_letters(line)}{row}")
    return prefix + (text[0] if len(set(map(tuple, cells))) == 1 else ":".join(text))


def _number(letters: str) -> int:
    total = 0
    for letter in letters:
        total = total * 26 + ord(letter) - 64
    return total


def _letters(number: int) -> str:
    letters = ""
    while number:
        number, remainder = divmod(number - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters


def _open(files: Path, fixture: str, sheet: str) -> tuple[Workbook, SheetPackage]:
    path = copy_fixture(files, fixture)
    workbook = load_workbook(path)
    package.capture(path, workbook, 50_000_000)
    return workbook, state_of(workbook).sheet(workbook[sheet])


def _sparklines(sheet: SheetPackage) -> dict[str, str]:
    xml = sheet.extensions[ext.SPARKLINES]
    found = re.findall(r"<x14:sparkline>(?:<xm:f>(.*?)</xm:f>)?<xm:sqref>(.*?)</xm:sqref>", xml)
    return {location: data for data, location in found}


def _corners(xml: str) -> list[tuple[int, int, int, int]]:
    numbers = re.findall(r"<(?:xdr:)?(col|colOff|row|rowOff)>(-?\d+)<", xml)
    values = [int(n) for _, n in numbers]
    return [tuple(values[i : i + 4]) for i in range(0, len(values), 4)]  # type: ignore[misc]


def _edit(sheet: str, axis: str, at: int, count: int = 1, *, delete: bool = False) -> LineEdit:
    return LineEdit(sheet, axis, at, count, delete)  # type: ignore[arg-type]


def _run(workbook: Workbook, edit: LineEdit) -> None:
    rewrite_lines(workbook, edit, _formula(edit))


def test_ranges_end_where_excel_puts_them() -> None:
    insert = _edit("Data", "rows", 3, 2)
    assert str(insert.area(_area("B2:B7"))) == "B2:B9"
    assert str(insert.area(_area("B4:B7"))) == "B6:B9"
    assert str(_edit("Data", "rows", 8).area(_area("B2:B7"))) == "B2:B8"  # grows over the row
    delete = _edit("Data", "rows", 4, 2, delete=True)
    assert str(delete.area(_area("B2:B7"))) == "B2:B5"
    assert str(delete.area(_area("B4:B5")) or "gone") == "gone"
    assert str(delete.area(_area("B5:B9"))) == "B4:B7"


def _area(text: str):
    from excel_mcp.refs import parse_range

    return parse_range(text)


def test_sparklines_gain_copies_in_inserted_rows_like_excel(files: Path) -> None:
    workbook, sheet = _open(files, "excel_sparklines.xlsx", "Data")

    _run(workbook, _edit("Data", "rows", 3))

    found = _sparklines(sheet)
    assert len(found) == 21
    assert found["G2"] == "Data!B2:F2"
    assert found["G3"] == "Data!B3:F3"  # the new row reads its own line
    assert found["G4"] == "Data!B4:F4"  # the old row 3 moved down with its data
    assert found["I8"] == "Data!B8:F8"
    rules = sheet.extensions[ext.CONDITIONAL_FORMATS]
    assert re.findall(r"<xm:sqref>(\w+:\w+)</xm:sqref>", rules) == ["B2:B8", "C2:C8"]


def test_sparklines_grow_below_the_last_row_but_not_above_the_first(files: Path) -> None:
    workbook, sheet = _open(files, "excel_sparklines.xlsx", "Data")
    _run(workbook, _edit("Data", "rows", 8))
    assert len(_sparklines(sheet)) == 21 and "G8" in _sparklines(sheet)

    workbook, sheet = _open(files, "excel_sparklines.xlsx", "Data")
    _run(workbook, _edit("Data", "rows", 2))
    found = _sparklines(sheet)
    assert len(found) == 18 and "G2" not in found and found["G3"] == "Data!B3:F3"


def test_sparklines_gain_copies_to_the_right_of_a_column(files: Path) -> None:
    workbook, sheet = _open(files, "excel_sparklines.xlsx", "Data")

    _run(workbook, _edit("Data", "columns", 10))

    found = _sparklines(sheet)
    assert found["J2"] == "Data!C2:G2"  # copied from I2 and moved one column right
    assert found["I2"] == "Data!B2:F2"


def test_deleted_columns_shrink_or_empty_the_data_of_sparklines(files: Path) -> None:
    workbook, sheet = _open(files, "excel_sparklines.xlsx", "Data")
    _run(workbook, _edit("Data", "columns", 3, delete=True))
    found = _sparklines(sheet)
    assert found["F2"] == "Data!B2:E2" and "I2" not in found

    workbook, sheet = _open(files, "excel_sparklines.xlsx", "Data")
    _run(workbook, _edit("Data", "columns", 2, 5, delete=True))
    found = _sparklines(sheet)
    assert set(found.values()) == {""} and "B2" in found  # Excel keeps them without data


def test_deleted_rows_take_their_sparklines(files: Path) -> None:
    workbook, sheet = _open(files, "excel_sparklines.xlsx", "Data")
    _run(workbook, _edit("Data", "rows", 4, 2, delete=True))
    found = _sparklines(sheet)
    assert len(found) == 12 and found["G4"] == "Data!B4:F4"


def test_other_sheets_are_not_moved_but_their_formulas_are_offered(files: Path) -> None:
    workbook, sheet = _open(files, "excel_sparklines.xlsx", "Data")
    before = sheet.extensions[ext.SPARKLINES]
    offered: list[tuple[str, str]] = []

    def formula(text: str, host: str) -> str:
        offered.append((text, host))
        return text

    rewrite_lines(workbook, _edit("Lists", "rows", 1), formula)

    assert sheet.extensions[ext.SPARKLINES] == before
    assert ("Data!B2:F2", "Data") in offered and ("Lists!$A$1:$A$3", "Data") in offered


def test_renaming_a_sheet_only_rewrites_formulas(files: Path) -> None:
    workbook, sheet = _open(files, "excel_sparklines.xlsx", "Data")

    rewrite_formulas(workbook, lambda text, host: text.replace("Lists!", "Items!"))

    assert "<xm:f>Items!$A$1:$A$3</xm:f>" in sheet.extensions[ext.DATA_VALIDATIONS]
    assert len(_sparklines(sheet)) == 18


def test_a_formula_with_xml_characters_survives_the_round_trip() -> None:
    xml = (
        f'<ext uri="{ext.DATA_VALIDATIONS}"><x14:dataValidations count="1">'
        '<x14:dataValidation type="custom"><x14:formula1><xm:f>A1&lt;&gt;&quot;a&amp;b&quot;'
        "</xm:f></x14:formula1><xm:sqref>B2</xm:sqref></x14:dataValidation>"
        "</x14:dataValidations></ext>"
    )
    found = {ext.DATA_VALIDATIONS: xml}
    seen: list[str] = []

    class Plain:
        edit: LineEdit | None = None

        def ranges(self, text: str, *, grow: bool = True) -> str | None:
            return text

        def formula(self, text: str) -> str:
            seen.append(text)
            return text

    ext.rewrite_extensions(found, Plain())

    assert seen == ['A1<>"a&b"']
    assert found[ext.DATA_VALIDATIONS] == xml.replace("&quot;", '"')


def test_shapes_keep_their_size_when_excel_keeps_it(files: Path) -> None:
    workbook, sheet = _open(files, "excel_controls.xlsx", "Data")

    _run(workbook, _edit("Data", "rows", 4, 2, delete=True))

    first, second = (_corners(a) for a in _control_anchors(sheet))
    assert first == [(2, 47625, 3, 0), (4, 95250, 4, 66675)]
    assert second == [(2, 47625, 4, 123825), (4, 95250, 6, 0)]


def _control_anchors(sheet: SheetPackage) -> list[str]:
    controls = next(xml for name, xml in sheet.elements if name == "controls")
    return controls.split("</control>")[:2]


def test_the_vml_of_form_controls_and_their_properties_follow(files: Path) -> None:
    workbook, sheet = _open(files, "excel_controls.xlsx", "Data")

    _run(workbook, _edit("Data", "rows", 4, 2, delete=True))

    anchors = [re.search(r"<x:Anchor>([^<]*)</x:Anchor>", s)[1] for s in sheet.vml if "Anchor" in s]  # type: ignore[index]
    assert anchors == ["2, 5, 3, 0, 4, 10, 4, 7", "2, 5, 4, 13, 4, 10, 6, 0"]
    parts = [link.target for link in sheet.links if not isinstance(link.target, str)]
    properties = " ".join(p.data.decode() for p in parts if b"formControlPr" in p.data)  # type: ignore[union-attr]
    assert 'fmlaLink="B1"' in properties and 'fmlaRange="A4:A5"' in properties


def test_inserted_rows_push_shapes_down(files: Path) -> None:
    workbook, sheet = _open(files, "excel_controls.xlsx", "Data")

    _run(workbook, _edit("Data", "rows", 1))

    first = _corners(_control_anchors(sheet)[0])
    assert first == [(2, 47625, 4, 57150), (4, 95250, 5, 123825)]
    assert re.search(r"<x:Anchor>\s*2, 5, 4, 6, 4, 10, 5, 13</x:Anchor>", "".join(sheet.vml))


@pytest.mark.parametrize(
    ("edit", "expected"),
    [
        (_edit("Data", "rows", 14, 7, delete=True), [(6, 152400, 13, 0), (11, 279400, 16, 63500)]),
        (
            _edit("Data", "columns", 1, 3, delete=True),
            [(3, 152400, 13, 63500), (8, 279400, 23, 63500)],
        ),
        (_edit("Data", "rows", 23, 5, delete=True), [(6, 152400, 13, 63500), (11, 279400, 22, 0)]),
        (_edit("Data", "rows", 19), [(6, 152400, 13, 63500), (11, 279400, 24, 63500)]),
    ],
)
def test_two_cell_shapes_move_and_resize_like_excel(
    files: Path, edit: LineEdit, expected: list[tuple[int, int, int, int]]
) -> None:
    workbook, sheet = _open(files, "excel_chartex.xlsx", "Data")

    _run(workbook, edit)

    assert _corners(sheet.anchors[0])[:2] == expected


def test_one_cell_shapes_in_deleted_rows_keep_their_size(files: Path) -> None:
    workbook, sheet = _open(files, "excel_controls.xlsx", "Data")

    _run(workbook, _edit("Data", "rows", 4, 9, delete=True))

    for anchor in _control_anchors(sheet):
        assert _corners(anchor) == [(2, 47625, 3, 0), (4, 95250, 4, 66675)]
