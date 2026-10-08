"""Keeping preserved content pointing at the right cells when the sheet is edited.

Preserved content holds ranges, formulas and positions of its own: sparklines, Excel 2010
conditional formats and validations, shapes and their anchors, form controls, ignored
errors, protected ranges, cell watches, sort state, and the data of newer charts. Code that
inserts or deletes rows and columns, or renames a sheet, calls these functions once so that
all of it follows, as it does in Excel.

``formula(text, host)`` is the caller's own rewriting of a formula: ``text`` has no leading
"=", ``host`` is the title of the sheet the text is on (what its unqualified references
mean), and the result is the text with its references updated, with ``#REF!`` where what
they pointed to is gone. It is called for every formula of preserved content, and may leave
a formula alone. Everything that is not a formula (ranges of this sheet, the cells that
shapes are anchored to) is moved here, by the `LineEdit`.

Call these before the rows and columns of the sheet are moved: the size of a shape that
Excel keeps when it moves it is taken from the sheet as it was. Cells, comments and their
marks move with openpyxl's own objects and need nothing here. Not handled: the selections
and panes of custom sheet views, and consolidation sources.
"""

import re
from collections.abc import Callable
from xml.sax.saxutils import escape

from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from excel_mcp.package import extensions
from excel_mcp.package.anchors import Geometry, Moved, move_anchors, move_vml_shape
from excel_mcp.package.lines import LineEdit
from excel_mcp.package.model import Part, SheetPackage, state_of
from excel_mcp.package.scan import unescape
from excel_mcp.refs import parse_range

Formula = Callable[[str, str], str]

_CONTROL_PROPERTIES = "application/vnd.ms-excel.controlproperties+xml"
_CHART_EX = "application/vnd.ms-office.chartex+xml"
_CONTROL_FORMULAS = re.compile(r'\b(fmla(?:Link|Range|Group|Txbx))="([^"]*)"')
_CHART_EX_FORMULA = re.compile(r"(<(?:\w+:)?f\b[^>]*>)([^<]*)(</(?:\w+:)?f>)")
_VML_FORMULA = re.compile(r"(<x:(Fmla(?:Link|Range|Txbx|Group))>)([^<]*)(</x:\2>)")
_RANGED_ITEMS = {
    "protectedRanges": ("protectedRange", "sqref"),
    "ignoredErrors": ("ignoredError", "sqref"),
    "cellWatches": ("cellWatch", "r"),
}


class _Rewriter:
    """The `extensions.Rewriter` of one sheet's preserved content."""

    def __init__(self, host: str, edit: LineEdit | None, formula: Formula) -> None:
        moves = edit is not None and host.casefold() == edit.sheet.casefold()
        self.edit = edit if moves else None
        self.host = host
        self._formula = formula

    def formula(self, text: str) -> str:
        return self._formula(text, self.host)

    def ranges(self, text: str, *, grow: bool = True) -> str | None:
        if self.edit is None:
            return text
        moved = [self.edit.area(parse_range(token), grow=grow) for token in text.split()]
        return " ".join(str(area) for area in moved if area is not None) or None


def rewrite_lines(workbook: Workbook, edit: LineEdit, formula: Formula) -> None:
    """Update the preserved content of every sheet for rows or columns that are inserted
    or deleted on ``edit.sheet``."""
    _rewrite(workbook, edit, formula)


def rewrite_formulas(workbook: Workbook, formula: Formula) -> None:
    """Update the formulas of preserved content only, as when a sheet is renamed."""
    _rewrite(workbook, None, formula)


def _rewrite(workbook: Workbook, edit: LineEdit | None, formula: Formula) -> None:
    for sheet, package in state_of(workbook).sheets.items():
        if isinstance(sheet, Worksheet):
            rewriter = _Rewriter(sheet.title, edit, formula)
            _sheet(package, rewriter, Geometry(sheet) if rewriter.edit else None)


def _sheet(package: SheetPackage, rewriter: _Rewriter, geometry: Geometry | None) -> None:
    extensions.rewrite_extensions(package.extensions, rewriter)
    package.rule_extensions = _rule_extensions(package.rule_extensions, rewriter)
    moved: Moved = {}
    edit = rewriter.edit
    if edit is not None and geometry is not None:
        package.anchors = [move_anchors(a, edit, geometry, moved) for a in package.anchors]
    elements = [(n, _element(n, x, rewriter, geometry, moved)) for n, x in package.elements]
    package.elements = [(name, xml) for name, xml in elements if xml is not None]
    package.vml = [_vml_shape(shape, rewriter, moved) for shape in package.vml]
    links = package.links + package.drawing_links + package.vml_links
    parts = {id(link.target): link.target for link in links if isinstance(link.target, Part)}
    for part in parts.values():
        _part(part, rewriter)


def _rule_extensions(
    found: dict[tuple[str, str], str], rewriter: _Rewriter
) -> dict[tuple[str, str], str]:
    """Rules are found again by their range, which moves like the range of their Excel 2010 half."""
    moved = {(rewriter.ranges(sqref), priority): xml for (sqref, priority), xml in found.items()}
    return {(sqref, priority): xml for (sqref, priority), xml in moved.items() if sqref is not None}


def _element(
    name: str, xml: str, rewriter: _Rewriter, geometry: Geometry | None, moved: Moved
) -> str | None:
    edit = rewriter.edit
    if edit is None:
        return xml
    if name in ("controls", "oleObjects") and geometry is not None:
        return move_anchors(xml, edit, geometry, moved)
    if name in _RANGED_ITEMS:
        return _ranged_items(xml, *_RANGED_ITEMS[name], rewriter)
    if name == "sortState":
        return _sort_state(xml, rewriter)
    return xml


def _ranged_items(xml: str, item: str, attribute: str, rewriter: _Rewriter) -> str | None:
    pattern = re.compile(rf"<(?P<p>(?:\w+:)?){item}\b[^>]*?(?:/>|>.*?</(?P=p){item}>)", re.S)

    def move(match: re.Match[str]) -> str:
        text = match[0]
        value = re.search(rf'\b{attribute}="([^"]*)"', text)
        assert value is not None
        moved = rewriter.ranges(value[1], grow=attribute == "sqref")
        return "" if moved is None else text.replace(value[0], f'{attribute}="{moved}"', 1)

    xml = pattern.sub(move, xml)
    return xml if pattern.search(xml) else None


def _sort_state(xml: str, rewriter: _Rewriter) -> str | None:
    conditions = re.compile(r"<(?:\w+:)?sortCondition\b[^>]*/>")
    ref = re.compile(r'\bref="([^"]*)"')

    def move(match: re.Match[str]) -> str:
        found = ref.search(match[0])
        assert found is not None
        moved = rewriter.ranges(found[1])
        return "" if moved is None else match[0].replace(found[0], f'ref="{moved}"', 1)

    before = len(conditions.findall(xml))
    head = re.match(r"<[^>]*>", xml)
    assert head is not None
    state = ref.search(head[0])
    if state is not None and rewriter.ranges(state[1]) is None:
        return None
    xml = conditions.sub(move, xml)
    if before and not conditions.search(xml):
        return None
    if state is not None:
        moved = rewriter.ranges(state[1])
        xml = xml.replace(head[0], head[0].replace(state[0], f'ref="{moved}"', 1), 1)
    return xml


def _vml_shape(shape: str, rewriter: _Rewriter, moved: Moved) -> str:
    if rewriter.edit is not None:
        shape = move_vml_shape(shape, rewriter.edit, moved)

    def formula(match: re.Match[str]) -> str:
        text = escape(rewriter.formula(unescape(match[3])))
        return f"{match[1]}{text}{match[4]}"

    return _VML_FORMULA.sub(formula, shape)


def _part(part: Part, rewriter: _Rewriter) -> None:
    """Form control properties and the cell references of newer charts, which live in parts."""
    if part.content_type == _CONTROL_PROPERTIES:
        pattern, rewrite = _CONTROL_FORMULAS, _attribute_formula
    elif part.content_type == _CHART_EX:
        pattern, rewrite = _CHART_EX_FORMULA, _text_formula
    else:
        return
    text = part.data.decode("utf-8")
    part.data = pattern.sub(lambda m: rewrite(m, rewriter), text).encode("utf-8")


def _attribute_formula(match: re.Match[str], rewriter: _Rewriter) -> str:
    text = rewriter.formula(unescape(match[2]))
    return f'{match[1]}="{escape(text, {chr(34): "&quot;"})}"'


def _text_formula(match: re.Match[str], rewriter: _Rewriter) -> str:
    return f"{match[1]}{escape(rewriter.formula(unescape(match[2])))}{match[3]}"


__all__ = ["Formula", "LineEdit", "rewrite_formulas", "rewrite_lines"]
