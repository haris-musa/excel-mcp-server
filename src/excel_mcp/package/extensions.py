"""The extensions Excel keeps in a sheet's ``<extLst>``: what is in them and how to edit it.

Sparklines, conditional formats and data validations of Excel 2010 and later live here, not
in the sheet proper. They hold ranges and formulas that must follow the cells they refer to.
"""

import re
import uuid
from collections.abc import Callable
from typing import Protocol
from xml.sax.saxutils import escape

from openpyxl.utils.cell import column_index_from_string, get_column_letter

from excel_mcp.package.lines import LineEdit
from excel_mcp.package.scan import unescape
from excel_mcp.refs import CellRange, clip_areas, parse_range

SPARKLINES = "{05C60535-1F16-4fd2-B633-F4F36F0B64E0}"
CONDITIONAL_FORMATS = "{78C0D931-6437-407d-A8EE-F0AAD7539E65}"
DATA_VALIDATIONS = "{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}"
# The <ext> inside a classic conditional format rule that names its Excel 2010 counterpart.
RULE_ID = "{B025F937-C7B1-47D3-B67F-A62EFF666E3E}"


class Rewriter(Protocol):
    """What the rewriting of references needs to know; see `package.references`."""

    edit: LineEdit | None

    def ranges(self, text: str, *, grow: bool = True) -> str | None:
        """A list of ranges of this sheet moved by the edit, None when nothing is left.

        Ranges that end just above inserted lines grow over them unless ``grow`` is False.
        """
        ...

    def formula(self, text: str) -> str:
        """A formula or operand, without the leading "=", with its references updated."""
        ...


_SPARKLINE = re.compile(r"<x14:sparkline>.*?</x14:sparkline>", re.S)
_EMPTY_GROUP = re.compile(
    r"<x14:sparklineGroup\b[^>]*>(?:(?!<x14:sparkline>).)*?</x14:sparklineGroup>", re.S
)
_RULE_GROUP = re.compile(r"<x14:conditionalFormatting\b[^>]*>.*?</x14:conditionalFormatting>", re.S)
_VALIDATION = re.compile(r"<x14:dataValidation\b[^>]*>.*?</x14:dataValidation>", re.S)
_DATED_GROUP = re.compile(
    r"(<x14:sparklineGroup(?!s)[^>]*>(?:(?!<x14:sparklines>).)*?)<xm:f>([^<]*)</xm:f>(<x14:sparklines>)",
    re.S,
)
_GROUP_DATES = re.compile(r"<xm:f>([^<]*)</xm:f>(?=<x14:sparklines>)")
_RANGE = re.compile(r"<xm:sqref>([^<]*)</xm:sqref>")
_FORMULA = re.compile(r"<xm:f>([^<]*)</xm:f>")
_FORMULA_ITEM = re.compile(
    r"<x14:cfRule\b.*?</x14:cfRule>|<x14:dataValidation\b.*?</x14:dataValidation>", re.S
)
_GUID = re.compile(r'(<x14:id>|\b(?:xr2:uid|id)=")(\{[0-9A-Fa-f-]{36}\})')
_CELL = re.compile(r"(\$?)\b([A-Z]{1,3})(\$?)(\d+)\b(?![!(\w])")


def rewrite_extensions(extensions: dict[str, str], rewriter: Rewriter) -> None:
    """Move the ranges and update the formulas of the sparklines, conditional formats and
    validations in a sheet's ``<ext>`` entries (`SheetPackage.extensions`), as Excel does.

    A sparkline whose location is deleted is deleted; one whose data is deleted loses its
    data. After an insertion just below or to the right of a sparkline, the inserted lines
    get copies of it that read the data of their own line. A conditional format or
    validation is deleted when its range is, and a formula it holds follows the edit.
    """
    for uri, xml in list(extensions.items()):
        if uri == SPARKLINES:
            kept = _sparklines(xml, lambda item: _moved_sparkline(item, rewriter))
            if kept is not None:
                kept = _GROUP_DATES.sub(lambda m: _moved_formula(m, rewriter), kept)
        elif uri == CONDITIONAL_FORMATS:
            kept = _items(xml, _RULE_GROUP, lambda item: _moved_item(item, rewriter), "<x14:cfRule")
        elif uri == DATA_VALIDATIONS:
            kept = _items(
                xml, _VALIDATION, lambda item: _moved_item(item, rewriter), "<x14:dataValidation "
            )
        else:
            continue
        if kept is None:
            del extensions[uri]
        else:
            extensions[uri] = kept


def clip_rules(extensions: dict[str, str], uri: str, hole: CellRange) -> None:
    """Take the cells of ``hole`` out of the ranges of the conditional formats (``uri`` is
    CONDITIONAL_FORMATS) or validations (DATA_VALIDATIONS) in a sheet's ``<ext>`` entries; one
    left without cells is deleted."""
    if uri not in extensions:
        return

    def clipped(item: str) -> str | None:
        location = _RANGE.search(item)
        areas = [parse_range(part) for part in location[1].split()] if location else []
        found = clip_areas(areas, hole)
        if found is None:
            return None
        kept, rows, columns = found
        item = _RANGE.sub(lambda _: f"<xm:sqref>{' '.join(map(str, kept))}</xm:sqref>", item)
        return _FORMULA.sub(lambda m: f"<xm:f>{_shifted(m[1], rows, columns)}</xm:f>", item)

    pattern, marker = (
        (_RULE_GROUP, "<x14:cfRule")
        if uri == CONDITIONAL_FORMATS
        else (_VALIDATION, "<x14:dataValidation ")
    )
    kept_xml = _items(extensions[uri], pattern, clipped, marker)
    if kept_xml is None:
        del extensions[uri]
    else:
        extensions[uri] = kept_xml


def forget_sheets(extensions: dict[str, str], names: set[str]) -> None:
    """Update what reads the deleted sheets ``names``, as Excel does: a sparkline loses its
    data, a group loses its date axis (Excel writes ``#REF!`` for the dates, a file that it
    cannot open again), and the formulas of validations and conditional formats become
    ``#REF!``."""

    def gone(text: str) -> bool:
        return any(refers_to(text, name) for name in names)

    def sparkline(item: str) -> list[str]:
        source = _FORMULA.search(item)
        return [item.replace(source[0], "")] if source and gone(unescape(source[1])) else [item]

    def group(match: re.Match[str]) -> str:
        if not gone(unescape(match[2])):
            return match[0]
        return match[1].replace(' dateAxis="1"', "").replace(' dateAxis="true"', "") + match[3]

    def rule(item: str) -> str:
        return _FORMULA.sub(lambda m: "<xm:f>#REF!</xm:f>" if gone(unescape(m[1])) else m[0], item)

    for uri, xml in list(extensions.items()):
        if uri == SPARKLINES:
            kept = _sparklines(xml, sparkline)
            if kept is not None:
                kept = _DATED_GROUP.sub(group, kept)
        elif uri in (CONDITIONAL_FORMATS, DATA_VALIDATIONS):
            kept = rule_sub(xml, rule)
        else:
            continue
        if kept is None:
            del extensions[uri]
        else:
            extensions[uri] = kept


def new_guid() -> str:
    return "{" + str(uuid.uuid4()).upper() + "}"


def edit_sparklines(extensions: dict[str, str], handle: Callable[[str], list[str]]) -> None:
    """Replace each ``<x14:sparkline>`` of the extensions by what ``handle`` returns for it
    (nothing deletes it); a group left without sparklines goes too."""
    xml = extensions.get(SPARKLINES)
    kept = None if xml is None else _sparklines(xml, handle)
    if kept is None:
        extensions.pop(SPARKLINES, None)
    else:
        extensions[SPARKLINES] = kept


def copied(
    extensions: dict[str, str],
    rule_extensions: dict[tuple[str, str], str],
    formula: Callable[[str], str],
) -> tuple[dict[str, str], dict[tuple[str, str], str]]:
    """The sparklines and conditional formats of a sheet as its copy has them: formulas
    rewritten by ``formula`` (no leading "="), and new ids as Excel gives them."""
    ids: dict[str, str] = {}

    def renewed(xml: str) -> str:
        return _GUID.sub(lambda m: m[1] + ids.setdefault(m[2], new_guid()), xml)

    def rewritten(xml: str) -> str:
        return _FORMULA.sub(lambda m: f"<xm:f>{escape(formula(unescape(m[1])))}</xm:f>", xml)

    kept = {
        uri: renewed(rewritten(xml))
        for uri, xml in extensions.items()
        if uri in (SPARKLINES, CONDITIONAL_FORMATS)
    }
    return kept, {key: renewed(xml) for key, xml in rule_extensions.items()}


def rule_sub(xml: str, rule: Callable[[str], str]) -> str:
    return _FORMULA_ITEM.sub(lambda m: rule(m[0]), xml)


def without_rules(extensions: dict[str, str], identifiers: set[str]) -> None:
    """Remove the Excel 2010 conditional format rules with these ids (``{GUID}``)."""
    xml = extensions.get(CONDITIONAL_FORMATS)
    if xml is None:
        return
    for identifier in identifiers:
        pattern = (
            rf'<x14:cfRule\b(?=[^>]*\bid="{re.escape(identifier)}")'
            r"(?:[^>]*/>|[^>]*>.*?</x14:cfRule>)"
        )
        xml = re.sub(pattern, "", xml, flags=re.S)
    emptied = (
        r"<x14:conditionalFormatting\b[^>]*>(?:(?!<x14:cfRule\b).)*?"
        r"</x14:conditionalFormatting>"
    )
    xml = re.sub(emptied, "", xml, flags=re.S)
    if "<x14:cfRule" in xml:
        extensions[CONDITIONAL_FORMATS] = xml
    else:
        del extensions[CONDITIONAL_FORMATS]


def refers_to(formula: str, sheet: str) -> bool:
    quoted = sheet.replace("'", "''")
    pattern = rf"(?<![\w.'])(?:'{re.escape(quoted)}'|{re.escape(sheet)})!"
    return re.search(pattern, formula, re.IGNORECASE) is not None


def _sparklines(xml: str, handle: Callable[[str], list[str]]) -> str | None:
    xml = _SPARKLINE.sub(lambda match: "".join(handle(match.group(0))), xml)
    xml = _EMPTY_GROUP.sub("", xml)
    return xml if re.search(r"<x14:sparklineGroup\b", xml) else None


def _moved_sparkline(item: str, rewriter: Rewriter) -> list[str]:
    location = _RANGE.search(item)
    source = _FORMULA.search(item)
    assert location is not None
    moved = rewriter.ranges(location[1], grow=False)
    if moved is None:
        return []
    item = _RANGE.sub(lambda _: f"<xm:sqref>{moved}</xm:sqref>", item)
    if source is not None:
        read = rewriter.formula(unescape(source[1]))
        if "#REF!" in read and "#REF!" not in source[1]:
            item = item.replace(source[0], "")
        else:
            item = item.replace(source[0], f"<xm:f>{escape(read)}</xm:f>")
    return [item, *_inserted_copies(item, location[1], rewriter)]


def _moved_formula(match: re.Match[str], rewriter: Rewriter) -> str:
    return f"<xm:f>{escape(rewriter.formula(unescape(match[1])))}</xm:f>"


def _inserted_copies(item: str, old: str, rewriter: Rewriter) -> list[str]:
    """Excel fills new lines below or to the right of a sparkline with copies of it."""
    edit = rewriter.edit
    if edit is None or edit.delete:
        return []
    cell = parse_range(old)
    line = cell.min_row if edit.axis == "rows" else cell.min_col
    if line != edit.at - 1:
        return []
    steps = range(1, edit.count + 1)
    return [_copy(item, *((n, 0) if edit.axis == "rows" else (0, n))) for n in steps]


def _copy(item: str, rows: int, columns: int) -> str:
    item = _RANGE.sub(lambda m: f"<xm:sqref>{_shifted(m[1], rows, columns)}</xm:sqref>", item)
    return _FORMULA.sub(lambda m: f"<xm:f>{_shifted(m[1], rows, columns)}</xm:f>", item)


def _shifted(text: str, rows: int, columns: int) -> str:
    """The relative cell references of ``text`` moved, as when the text is copied."""

    def move(match: re.Match[str]) -> str:
        column_fixed, letters, row_fixed, number = match.groups()
        column = column_index_from_string(letters)
        column += 0 if column_fixed else columns
        row = int(number) + (0 if row_fixed else rows)
        return f"{column_fixed}{get_column_letter(column)}{row_fixed}{row}"

    return _CELL.sub(move, text)


def _moved_item(item: str, rewriter: Rewriter) -> str | None:
    location = _RANGE.search(item)
    moved = rewriter.ranges(location[1]) if location else None
    if moved is None:
        return None
    item = _RANGE.sub(lambda _: f"<xm:sqref>{moved}</xm:sqref>", item)
    return _FORMULA.sub(lambda m: f"<xm:f>{escape(rewriter.formula(unescape(m[1])))}</xm:f>", item)


def _items(
    xml: str, pattern: re.Pattern[str], handle: Callable[[str], str | None], marker: str
) -> str | None:
    xml = pattern.sub(lambda match: handle(match.group(0)) or "", xml)
    if marker not in xml:
        return None
    if marker == "<x14:dataValidation ":
        count = xml.count(marker)
        counted = r'(<x14:dataValidations\b[^>]*?\bcount=")\d+(")'
        xml = re.sub(counted, lambda m: f"{m[1]}{count}{m[2]}", xml, count=1)
    return xml
