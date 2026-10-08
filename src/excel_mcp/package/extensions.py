"""The extensions Excel keeps in a sheet's ``<extLst>``: what is in them and how to edit it.

Sparklines, conditional formats and data validations of Excel 2010 and later live here, not
in the sheet proper. They hold ranges and formulas that must follow the cells they refer to.
"""

import re
from collections.abc import Callable
from typing import Literal

SPARKLINES = "{05C60535-1F16-4fd2-B633-F4F36F0B64E0}"
CONDITIONAL_FORMATS = "{78C0D931-6437-407d-A8EE-F0AAD7539E65}"
DATA_VALIDATIONS = "{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}"

Kind = Literal["range", "formula"]
Update = Callable[[str, Kind], str | None]

_SPARKLINE = re.compile(r"<x14:sparkline>.*?</x14:sparkline>", re.S)
_EMPTY_GROUP = re.compile(
    r"<x14:sparklineGroup\b[^>]*>(?:(?!<x14:sparkline>).)*?</x14:sparklineGroup>", re.S
)
_RULE_GROUP = re.compile(r"<x14:conditionalFormatting\b[^>]*>.*?</x14:conditionalFormatting>", re.S)
_VALIDATION = re.compile(r"<x14:dataValidation\b[^>]*>.*?</x14:dataValidation>", re.S)
_RANGE = re.compile(r"<xm:sqref>([^<]*)</xm:sqref>")
_FORMULA = re.compile(r"<xm:f>([^<]*)</xm:f>")


def rewrite_references(extensions: dict[str, str], update: Update) -> None:
    """Pass the ranges and formulas of sparklines, conditional formats and validations on.

    ``extensions`` is the sheet's ``<ext>`` entries by uri (`SheetPackage.extensions`).
    ``update(text, kind)`` returns the new text, or None when what the text refers to is
    gone. As in Excel, a sparkline is deleted when its location or its data is gone; a
    conditional format or validation is deleted when its range is gone, and gets ``#REF!``
    for a formula that is gone.
    """
    for uri, xml in list(extensions.items()):
        if uri == SPARKLINES:
            kept = _sparklines(xml, update)
        elif uri == CONDITIONAL_FORMATS:
            kept = _items(xml, _RULE_GROUP, update, "<x14:cfRule")
        elif uri == DATA_VALIDATIONS:
            kept = _items(xml, _VALIDATION, update, "<x14:dataValidation ")
        else:
            continue
        if kept is None:
            del extensions[uri]
        else:
            extensions[uri] = kept


def forget_sheets(extensions: dict[str, str], names: set[str]) -> None:
    """Drop what reads the deleted sheets ``names``: Excel deletes such sparklines."""

    def update(text: str, kind: Kind) -> str | None:
        gone = kind == "formula" and any(refers_to(text, name) for name in names)
        return None if gone else text

    rewrite_references(extensions, update)


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


def _sparklines(xml: str, update: Update) -> str | None:
    def rewrite(match: re.Match[str]) -> str:
        item = match.group(0)
        location = _RANGE.search(item)
        source = _FORMULA.search(item)
        moved = update(location[1], "range") if location else None
        read = update(source[1], "formula") if source else None
        if moved is None or read is None:
            return ""
        item = _RANGE.sub(lambda _: f"<xm:sqref>{moved}</xm:sqref>", item)
        return _FORMULA.sub(lambda _: f"<xm:f>{read}</xm:f>", item)

    xml = _EMPTY_GROUP.sub("", _SPARKLINE.sub(rewrite, xml))
    return xml if re.search(r"<x14:sparklineGroup\b", xml) else None


def _items(xml: str, pattern: re.Pattern[str], update: Update, marker: str) -> str | None:
    def rewrite(match: re.Match[str]) -> str:
        item = match.group(0)
        location = _RANGE.search(item)
        moved = update(location[1], "range") if location else None
        if moved is None:
            return ""
        item = _RANGE.sub(lambda _: f"<xm:sqref>{moved}</xm:sqref>", item)
        return _FORMULA.sub(lambda m: f"<xm:f>{update(m[1], 'formula') or '#REF!'}</xm:f>", item)

    xml = pattern.sub(rewrite, xml)
    if marker not in xml:
        return None
    if marker == "<x14:dataValidation ":
        count = xml.count(marker)
        counted = r'(<x14:dataValidations\b[^>]*?\bcount=")\d+(")'
        xml = re.sub(counted, lambda m: f"{m[1]}{count}{m[2]}", xml, count=1)
    return xml
