"""Data bars and icon sets as Excel 2010 and later write them.

A data bar is a classic rule plus a rule in the sheet's Excel 2010 extension that holds what
the classic one cannot say: borders, negative bars, axis and direction. The classic rule names
its partner by id. Icon sets with custom icons, and the newer sets (stars, triangles, boxes),
exist in the extension only. The package layer carries the extension through edits.
"""

import re
from typing import Literal, cast, get_args
from xml.sax.saxutils import escape

from openpyxl.formatting.rule import DataBarRule, Rule
from openpyxl.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from excel_mcp.errors import InvalidArgumentError
from excel_mcp.formulas import storable_operand
from excel_mcp.operations.conditional_kinds import IconSetName, ScaleType, ThresholdType
from excel_mcp.operations.conditional_rule import ConditionalFormat
from excel_mcp.operations.formatting import parse_color
from excel_mcp.operations.x14_xml import X14, XM, attributes, color, number_text
from excel_mcp.package import state_of
from excel_mcp.package.extensions import CONDITIONAL_FORMATS, RULE_ID, new_guid

_NEWER_ICON_SETS = {"3Stars", "3Triangles", "5Boxes"}
_DIRECTIONS = {"context": None, "left_to_right": "leftToRight", "right_to_left": "rightToLeft"}
_AXES = {"automatic": None, "middle": "middle", "none": "none"}
_VALUED = {"number": "num", "percent": "percent", "percentile": "percentile"}
_PRIORITY = re.compile(r'(<x14:cfRule\b[^>]*?\bpriority=")(\d+)(")')
_BLOCK = re.compile(r"<x14:conditionalFormatting\b.*?</x14:conditionalFormatting>", re.S)
_STANDALONE = re.compile(r'<x14:cfRule\b(?=[^>]*\bpriority=")[^>]*?\btype="(\w+)"')
_SQREF = re.compile(r"<xm:sqref>([^<]*)</xm:sqref>")
_ICON = re.compile(r"(\w+):(\d)")


def is_extended(rule: ConditionalFormat) -> bool:
    """Whether the rule needs the Excel 2010 extension."""
    return rule.type == "data_bar" or (
        rule.type == "icon_set" and (rule.icon_set in _NEWER_ICON_SETS or rule.icons is not None)
    )


def icon_points(rule: ConditionalFormat) -> list[tuple[ThresholdType, float]]:
    """Where each icon of the set starts: (what the number means, the number)."""
    if rule.icon_set is None:
        raise InvalidArgumentError("icon_set rules need an icon_set.")
    icons = int(rule.icon_set[0])
    thresholds = rule.thresholds or [round(100 * index / icons) for index in range(1, icons)]
    if len(thresholds) != icons - 1:
        raise InvalidArgumentError(f"{rule.icon_set} needs {icons - 1} thresholds.")
    if thresholds != sorted(thresholds):
        raise InvalidArgumentError("thresholds must be in ascending order.")
    kinds: dict[str, ThresholdType] = {
        "percent": "percent",
        "number": "num",
        "percentile": "percentile",
    }
    points: list[tuple[ThresholdType, float]] = [("percent", 0)]
    return points + [(kinds[rule.threshold_type], value) for value in thresholds]


def claim_priority(sheet: Worksheet, requested: int | None) -> int:
    """The priority for a new rule: last, or ``requested`` with the rules from there moving
    down. Rules in the extension share the numbering with the classic ones."""
    package = state_of(cast(Workbook, sheet.parent)).sheet(sheet)
    entries = list(sheet.conditional_formatting)
    xml = package.extensions.get(CONDITIONAL_FORMATS, "")
    used = [rule.priority or 0 for entry in entries for rule in entry.rules]
    used += [int(found[1]) for found in _PRIORITY.findall(xml)]
    if requested is None:
        return max(used, default=0) + 1
    moved: dict[tuple[str, str], tuple[str, str]] = {}
    for entry in entries:
        for rule in entry.rules:
            if rule.priority and rule.priority >= requested:
                before = (str(entry.sqref), str(rule.priority))
                rule.priority += 1
                moved[before] = (before[0], str(rule.priority))
    package.rule_extensions = {moved.get(k, k): v for k, v in package.rule_extensions.items()}
    if xml:
        shifted = _PRIORITY.sub(
            lambda m: f"{m[1]}{int(m[2]) + (int(m[2]) >= requested)}{m[3]}", xml
        )
        package.extensions[CONDITIONAL_FORMATS] = shifted
    return requested


def add_extended(
    sheet: Worksheet,
    area: str,
    rule: ConditionalFormat,
    priority: int,
    names: list[str],
) -> Rule | None:
    """Write the extension half of a data bar or icon set. Returns the classic rule that
    goes with it, None for an icon set, which exists in the extension only."""
    package = state_of(cast(Workbook, sheet.parent)).sheet(sheet)
    guid = new_guid()
    classic = None
    if rule.type == "data_bar":
        classic, xml = _data_bar(rule, guid, names)
        package.rule_extensions[(area, str(priority))] = _link(guid)
    else:
        xml = _icon_set(rule, guid, priority)
    block = (
        f'<x14:conditionalFormatting xmlns:xm="{XM}">{xml}'
        f"<xm:sqref>{area}</xm:sqref></x14:conditionalFormatting>"
    )
    existing = package.extensions.get(CONDITIONAL_FORMATS)
    if existing is None:
        package.extensions[CONDITIONAL_FORMATS] = (
            f'<ext uri="{CONDITIONAL_FORMATS}" xmlns:x14="{X14}">'
            f"<x14:conditionalFormattings>{block}</x14:conditionalFormattings></ext>"
        )
    else:
        head, closing, tail = existing.rpartition("</x14:conditionalFormattings>")
        package.extensions[CONDITIONAL_FORMATS] = head + block + closing + tail
    return classic


def extended_rules(sheet: Worksheet) -> list[tuple[str, str]]:
    """(range, type) of the rules that exist in the extension only."""
    package = state_of(cast(Workbook, sheet.parent)).sheets.get(sheet)
    xml = package.extensions.get(CONDITIONAL_FORMATS, "") if package else ""
    found = []
    for block in _BLOCK.findall(xml):
        sqref = _SQREF.search(block)
        found += [(sqref[1] if sqref else "", kind) for kind in _STANDALONE.findall(block)]
    return found


def _link(guid: str) -> str:
    return f'<extLst><ext uri="{RULE_ID}" xmlns:x14="{X14}"><x14:id>{guid}</x14:id></ext></extLst>'


def _data_bar(rule: ConditionalFormat, guid: str, names: list[str]) -> tuple[Rule, str]:
    if len(rule.colors or []) != 1:
        raise InvalidArgumentError("data_bar needs exactly 1 color.")
    fill = parse_color((rule.colors or [""])[0])
    low = _point(rule.min_type, rule.min_value, "min", names)
    high = _point(rule.max_type, rule.max_value, "max", names)
    if rule.negative_border_color and not rule.border_color:
        raise InvalidArgumentError("negative_border_color needs a border_color.")
    classic = DataBarRule(
        start_type=low[0],
        start_value=low[2],
        end_type=high[0],
        end_value=high[2],
        color=fill,
        showValue=False if rule.hide_values else None,
    )
    axis_color = rule.axis_color or ("000000" if rule.axis != "none" else None)
    options = attributes(
        [
            ("minLength", "0"),
            ("maxLength", "100"),
            ("border", "1" if rule.border_color else None),
            ("gradient", "0" if rule.bar_fill == "solid" else None),
            ("direction", _DIRECTIONS[rule.bar_direction]),
            ("negativeBarColorSameAsPositive", None if rule.negative_color else "1"),
            ("negativeBarBorderColorSameAsPositive", "0" if rule.negative_border_color else None),
            ("axisPosition", _AXES[rule.axis]),
        ]
    )
    colors = "".join(
        color(tag, value)
        for tag, value in [
            ("borderColor", rule.border_color),
            ("negativeFillColor", rule.negative_color),
            ("negativeBorderColor", rule.negative_border_color),
            ("axisColor", axis_color),
        ]
        if value
    )
    points = "".join(_x14_point(kind, value) for _, kind, value in (low, high))
    xml = (
        f'<x14:cfRule type="dataBar" id="{guid}">'
        f"<x14:dataBar{options}>{points}{colors}</x14:dataBar></x14:cfRule>"
    )
    return classic, xml


def _point(
    kind: ScaleType, value: float | str | None, side: Literal["min", "max"], names: list[str]
) -> tuple[str, str, str | None]:
    """(classic type, extension type, value) of the end of a data bar's scale."""
    label = f"{side}_type"
    if kind in ("automatic", "lowest", "highest"):
        if kind != "automatic" and kind != ("lowest" if side == "min" else "highest"):
            raise InvalidArgumentError(f"{label} cannot be {kind!r}.")
        if value is not None:
            raise InvalidArgumentError(
                f"{side}_value is for {label} number, percent, percentile or formula."
            )
        return side, f"auto{side.capitalize()}" if kind == "automatic" else side, None
    if value is None:
        raise InvalidArgumentError(f"{label} {kind} needs {side}_value.")
    if kind == "formula":
        return "formula", "formula", storable_operand(str(value), names)
    number = _number(value, f"{side}_value")
    if kind != "number" and not 0 <= number <= 100:
        raise InvalidArgumentError(f"{side}_value is a {kind}: 0 to 100.")
    return _VALUED[kind], _VALUED[kind], number_text(number)


def _number(value: float | str, label: str) -> float:
    try:
        return float(value)
    except ValueError:
        raise InvalidArgumentError(f"{label} must be a number.") from None


def _x14_point(kind: str, value: str | None) -> str:
    if value is None:
        return f'<x14:cfvo type="{kind}"/>'
    return f'<x14:cfvo type="{kind}"><xm:f>{escape(value)}</xm:f></x14:cfvo>'


def _icon_set(rule: ConditionalFormat, guid: str, priority: int) -> str:
    points = icon_points(rule)
    custom = _custom_icons(rule, len(points))
    options = attributes(
        [
            ("iconSet", rule.icon_set),
            ("showValue", "0" if rule.hide_values else None),
            ("reverse", "1" if rule.reverse else None),
            ("custom", "1" if custom else None),
        ]
    )
    cfvo = "".join(_x14_point(kind, number_text(value)) for kind, value in points)
    stop = ' stopIfTrue="1"' if rule.stop_if_true else ""
    return (
        f'<x14:cfRule type="iconSet" priority="{priority}"{stop} id="{guid}">'
        f"<x14:iconSet{options}>{cfvo}{custom}</x14:iconSet></x14:cfRule>"
    )


def _custom_icons(rule: ConditionalFormat, count: int) -> str:
    if rule.icons is None:
        return ""
    if rule.reverse:
        raise InvalidArgumentError("icons already give the order; remove reverse.")
    if len(rule.icons) != count:
        raise InvalidArgumentError(f"{rule.icon_set} has {count} ranges; icons needs {count}.")
    sets = {name for part in get_args(IconSetName) for name in get_args(part)}
    shown = []
    for icon in rule.icons:
        if icon == "none":
            shown.append('<x14:cfIcon iconSet="NoIcons" iconId="0"/>')
            continue
        found = _ICON.fullmatch(icon)
        if not found or found[1] not in sets or not 1 <= int(found[2]) <= int(found[1][0]):
            raise InvalidArgumentError(
                f"Invalid icon {icon!r}. Use 'none' or a set and the icon's position in it, "
                "lowest first, such as '3Arrows:1'."
            )
        shown.append(f'<x14:cfIcon iconSet="{found[1]}" iconId="{int(found[2]) - 1}"/>')
    return "".join(shown)
