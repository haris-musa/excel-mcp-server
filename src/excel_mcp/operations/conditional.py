"""Conditional formatting rules of the kinds Excel's Conditional Formatting menu offers."""

from typing import cast

from openpyxl.formatting.rule import (
    CellIsRule,
    ColorScaleRule,
    FormatObject,
    FormulaRule,
    IconSet,
    Rule,
)
from openpyxl.styles import Font, PatternFill
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.worksheet.worksheet import Worksheet

from excel_mcp.errors import InvalidArgumentError
from excel_mcp.formulas import storable_operand
from excel_mcp.operations import conditional_x14 as x14
from excel_mcp.operations.conditional_kinds import (
    ALWAYS,
    CELL_FORMULAS,
    FIELDS,
    PERIOD_FORMULAS,
    TEXT_RULES,
    ClassicIconSet,
)
from excel_mcp.operations.conditional_rule import ConditionalFormat
from excel_mcp.operations.formatting import parse_color
from excel_mcp.refs import parse_range
from excel_mcp.workspace import sheet_names


def add_conditional_format(sheet: Worksheet, ref: str, rule: ConditionalFormat) -> str:
    area = parse_range(ref)
    rule.reject_unused(FIELDS[rule.type] | ALWAYS, f"A {rule.type} rule")
    names = sheet_names(sheet)
    priority = x14.claim_priority(sheet, rule.priority)
    if x14.is_extended(rule):
        built = x14.add_extended(sheet, str(area), rule, priority, names)
    else:
        built = _build_rule(rule, area.top_left, names)
    if built is not None:
        built.priority = priority
        if rule.stop_if_true:
            built.stopIfTrue = True
        sheet.conditional_formatting.add(str(area), built)
    return str(area)


def _build_rule(rule: ConditionalFormat, top_left: str, names: list[str]) -> Rule:
    match rule.type:
        case "color_scale":
            return _color_scale(rule)
        case "icon_set":
            return _icon_set(rule)
    cell = top_left
    match rule.type:
        case "cell_value":
            if rule.operator is None or not rule.values:
                raise InvalidArgumentError("cell_value needs an operator and values.")
            if any(not value.strip() for value in rule.values):
                raise InvalidArgumentError("cell_value values cannot be empty.")
            values = [storable_operand(value, names) for value in rule.values]
            dxf = _differential_style(rule)
            return CellIsRule(operator=rule.operator, formula=values, fill=dxf.fill, font=dxf.font)
        case "formula":
            if not rule.formula:
                raise InvalidArgumentError("formula rules need a formula.")
            formula = storable_operand(rule.formula, names)
            dxf = _differential_style(rule)
            return FormulaRule(formula=[formula], fill=dxf.fill, font=dxf.font)
        case "top" | "bottom":
            if rule.count is None or rule.count > (100 if rule.percent else 1000):
                raise InvalidArgumentError("top and bottom need a count (at most 100 if percent).")
            return Rule(
                type="top10",
                rank=rule.count,
                percent=rule.percent or None,
                bottom=rule.type == "bottom" or None,
                dxf=_differential_style(rule),
            )
        case "above_average" | "below_average":
            return Rule(
                type="aboveAverage",
                aboveAverage=None if rule.type == "above_average" else False,
                equalAverage=rule.include_equal or None,
                stdDev=rule.std_dev,
                dxf=_differential_style(rule),
            )
        case "duplicate" | "unique":
            kind = "duplicateValues" if rule.type == "duplicate" else "uniqueValues"
            return Rule(type=kind, dxf=_differential_style(rule))
        case "contains_text" | "not_contains_text" | "begins_with" | "ends_with":
            if not rule.text:
                raise InvalidArgumentError(f"{rule.type} needs text.")
            kind, operator, template = TEXT_RULES[rule.type]
            literal = '"' + rule.text.replace('"', '""') + '"'
            formula = storable_operand(template.format(t=literal, c=cell), names)
            return Rule(
                type=kind,
                operator=operator,
                text=rule.text,
                formula=[formula],
                dxf=_differential_style(rule),
            )
        case "date":
            if rule.period is None:
                raise InvalidArgumentError("date rules need a period.")
            formula = storable_operand(PERIOD_FORMULAS[rule.period].format(c=cell), names)
            return Rule(
                type="timePeriod",
                timePeriod=rule.period,
                formula=[formula],
                dxf=_differential_style(rule),
            )
        case _:
            kind, template = CELL_FORMULAS[rule.type]
            formula = storable_operand(template.format(c=cell), names)
            return Rule(type=kind, formula=[formula], dxf=_differential_style(rule))


def _differential_style(rule: ConditionalFormat) -> DifferentialStyle:
    if not (rule.fill_color or rule.font_color):
        raise InvalidArgumentError(f"{rule.type} rules need a fill_color or font_color.")
    fill = (
        PatternFill(fill_type="solid", bgColor=parse_color(rule.fill_color))
        if rule.fill_color
        else None
    )
    font = Font(color=parse_color(rule.font_color)) if rule.font_color else None
    return DifferentialStyle(fill=fill, font=font)


def _color_scale(rule: ConditionalFormat) -> Rule:
    colors = [parse_color(color) for color in rule.colors or []]
    if len(colors) == 2:
        return ColorScaleRule(
            start_type="min", start_color=colors[0], end_type="max", end_color=colors[1]
        )
    if len(colors) == 3:
        return ColorScaleRule(
            start_type="min",
            start_color=colors[0],
            mid_type="percentile",
            mid_value=50,
            mid_color=colors[1],
            end_type="max",
            end_color=colors[2],
        )
    raise InvalidArgumentError("color_scale needs 2 or 3 colors.")


def _icon_set(rule: ConditionalFormat) -> Rule:
    if rule.icon_set is None:
        raise InvalidArgumentError("icon_set rules need an icon_set.")
    points = [FormatObject(type=kind, val=_plain(value)) for kind, value in x14.icon_points(rule)]
    icons_rule = IconSet(
        iconSet=cast(ClassicIconSet, rule.icon_set),  # the newer sets are extended
        cfvo=points,
        showValue=False if rule.hide_values else None,
        reverse=rule.reverse or None,
    )
    return Rule(type="iconSet", iconSet=icons_rule)


def _plain(value: float) -> float:
    return int(value) if value == int(value) else value
