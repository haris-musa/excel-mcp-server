"""Conditional formatting rules of the kinds Excel's Conditional Formatting menu offers."""

from typing import Literal

from openpyxl.formatting.rule import (
    CellIsRule,
    ColorScaleRule,
    DataBarRule,
    FormatObject,
    FormulaRule,
    IconSet,
    Rule,
)
from openpyxl.styles import Font, PatternFill
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.worksheet.worksheet import Worksheet
from pydantic import Field

from excel_mcp.errors import InvalidArgumentError
from excel_mcp.formulas import storable_operand
from excel_mcp.inputs import InputModel
from excel_mcp.operations.conditional_kinds import (
    ALWAYS,
    CELL_FORMULAS,
    FIELDS,
    PERIOD_FORMULAS,
    TEXT_RULES,
    IconSetName,
    Period,
    RuleType,
    ThresholdType,
)
from excel_mcp.operations.formatting import parse_color
from excel_mcp.operations.rules import Operator
from excel_mcp.refs import parse_range
from excel_mcp.workspace import sheet_names


class ConditionalFormat(InputModel):
    """Only the fields named for the ``type`` apply; others are rejected."""

    type: RuleType = Field(
        description="top/bottom: highest or lowest `count` values (or percent). "
        "above_average/below_average: compared with the range's average. "
        "date: dates in a `period`."
    )
    colors: list[str] | None = Field(
        default=None, description="color_scale: 2 or 3, lowest to highest. data_bar: 1."
    )
    operator: Operator | None = Field(default=None, description="cell_value.")
    values: list[str] | None = Field(
        default=None,
        description="cell_value: 1, or 2 for between/notBetween. Numbers, quoted text such "
        "as '\"Done\"', or formulas.",
    )
    formula: str | None = Field(
        default=None,
        description="formula: true for highlighted cells, written for the range's top-left "
        "cell, e.g. '=$C2>100'.",
    )
    count: int | None = Field(
        default=None, ge=1, le=1000, description="top, bottom: how many (1-100 if percent)."
    )
    percent: bool = Field(default=False, description="top, bottom: `count` is a percentage.")
    std_dev: int | None = Field(
        default=None, ge=1, le=3, description="above/below_average: standard deviations."
    )
    include_equal: bool = Field(default=False, description="above/below_average.")
    text: str | None = Field(default=None, max_length=255, description="contains_text etc.")
    period: Period | None = Field(default=None, description="date.")
    icon_set: IconSetName | None = Field(default=None, description="icon_set.")
    thresholds: list[float] | None = Field(
        default=None,
        description="icon_set: where icons 2..n start, lowest first (n-1 values). "
        "Default: equal shares as in Excel.",
    )
    threshold_type: Literal["percent", "number", "percentile"] = Field(
        default="percent", description="icon_set: what `thresholds` mean."
    )
    reverse: bool = Field(default=False, description="icon_set: reverse the icon order.")
    icon_only: bool = Field(default=False, description="icon_set: hide the cell values.")
    fill_color: str | None = Field(default=None, description="Every type but the scales/icons.")
    font_color: str | None = Field(default=None, description="Like fill_color.")
    stop_if_true: bool = Field(default=False, description="Skip lower-priority rules if met.")
    priority: int | None = Field(
        default=None,
        ge=1,
        description="1 is evaluated first; rules at or below it move down. Default: last.",
    )


def add_conditional_format(sheet: Worksheet, ref: str, rule: ConditionalFormat) -> str:
    area = parse_range(ref)
    rule.reject_unused(FIELDS[rule.type] | ALWAYS, f"A {rule.type} rule")
    built = _build_rule(rule, area.top_left, sheet_names(sheet))
    if rule.stop_if_true:
        built.stopIfTrue = True
    formatting = sheet.conditional_formatting
    existing = [entry for ranges in formatting for entry in ranges.rules]
    built.priority = (
        rule.priority or max((entry.priority or 0 for entry in existing), default=0) + 1
    )
    if rule.priority:
        for entry in existing:
            if entry.priority and entry.priority >= rule.priority:
                entry.priority += 1
    formatting.add(str(area), built)
    return str(area)


def _build_rule(rule: ConditionalFormat, top_left: str, names: list[str]) -> Rule:
    match rule.type:
        case "color_scale":
            return _color_scale(rule)
        case "data_bar":
            if len(rule.colors or []) != 1:
                raise InvalidArgumentError("data_bar needs exactly 1 color.")
            color = parse_color((rule.colors or [""])[0])
            return DataBarRule(start_type="min", end_type="max", color=color)
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
    kind = kinds[rule.threshold_type]
    points = [FormatObject(type="percent", val=0)]
    points += [
        FormatObject(type=kind, val=int(value) if value == int(value) else value)
        for value in thresholds
    ]
    icons_rule = IconSet(
        iconSet=rule.icon_set,
        cfvo=points,
        showValue=False if rule.icon_only else None,
        reverse=rule.reverse or None,
    )
    return Rule(type="iconSet", iconSet=icons_rule)
