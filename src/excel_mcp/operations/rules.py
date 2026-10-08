"""Conditional formatting and data validation rules."""

from typing import Literal

from openpyxl.formatting.rule import CellIsRule, ColorScaleRule, DataBarRule, FormulaRule
from openpyxl.styles import Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.worksheet import Worksheet
from pydantic import Field

from excel_mcp.errors import InvalidArgumentError
from excel_mcp.formulas import storable_operand
from excel_mcp.inputs import InputModel
from excel_mcp.operations.formatting import parse_color
from excel_mcp.refs import parse_range
from excel_mcp.workspace import sheet_names

Operator = Literal[
    "between",
    "notBetween",
    "equal",
    "notEqual",
    "greaterThan",
    "lessThan",
    "greaterThanOrEqual",
    "lessThanOrEqual",
]


class ConditionalFormat(InputModel):
    """Which fields apply depends on ``type``."""

    type: Literal["color_scale", "data_bar", "cell_value", "formula"]
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
    fill_color: str | None = Field(default=None, description="cell_value, formula.")
    font_color: str | None = Field(default=None, description="cell_value, formula.")


class DataValidationRule(InputModel):
    """Which fields apply depends on ``type``."""

    type: Literal["list", "whole", "decimal", "date", "text_length", "custom"]
    options: list[str] | None = Field(default=None, description="list: allowed values.")
    operator: Operator | None = Field(
        default=None, description="whole, decimal, date, text_length."
    )
    minimum: str | None = Field(
        default=None, description="Number or formula; dates as 'DATE(2026,1,31)'."
    )
    maximum: str | None = Field(default=None, description="For between/notBetween.")
    formula: str | None = Field(
        default=None, description="custom: for the range's top-left cell, e.g. '=A2>B2'."
    )
    allow_blank: bool = True
    error_message: str | None = Field(default=None, description="Shown for rejected input.")
    prompt: str | None = Field(default=None, description="Shown when a cell is selected.")


def add_conditional_format(sheet: Worksheet, ref: str, rule: ConditionalFormat) -> str:
    target = str(parse_range(ref))
    sheet.conditional_formatting.add(target, _build_conditional_rule(rule, sheet_names(sheet)))
    return target


def _build_conditional_rule(rule: ConditionalFormat, names: list[str]):
    colors = [parse_color(color) for color in rule.colors or []]
    fill = (
        PatternFill(fill_type="solid", bgColor=parse_color(rule.fill_color))
        if rule.fill_color
        else None
    )
    font = Font(color=parse_color(rule.font_color)) if rule.font_color else None

    match rule.type:
        case "color_scale" if len(colors) == 2:
            return ColorScaleRule(
                start_type="min", start_color=colors[0], end_type="max", end_color=colors[1]
            )
        case "color_scale" if len(colors) == 3:
            return ColorScaleRule(
                start_type="min",
                start_color=colors[0],
                mid_type="percentile",
                mid_value=50,
                mid_color=colors[1],
                end_type="max",
                end_color=colors[2],
            )
        case "color_scale":
            raise InvalidArgumentError("color_scale needs 2 or 3 colors.")
        case "data_bar" if len(colors) == 1:
            return DataBarRule(start_type="min", end_type="max", color=colors[0])
        case "data_bar":
            raise InvalidArgumentError("data_bar needs exactly 1 color.")
        case "cell_value":
            if rule.operator is None or not rule.values:
                raise InvalidArgumentError("cell_value needs an operator and values.")
            if any(not value.strip() for value in rule.values):
                raise InvalidArgumentError("cell_value values cannot be empty.")
            values = [_checked_operand(value, names) for value in rule.values]
            return CellIsRule(operator=rule.operator, formula=values, fill=fill, font=font)
        case "formula":
            if not rule.formula:
                raise InvalidArgumentError("formula rules need a formula.")
            formula = storable_operand(rule.formula, names)
            return FormulaRule(formula=[formula], fill=fill, font=font)


def add_data_validation(sheet: Worksheet, ref: str, rule: DataValidationRule) -> str:
    target = str(parse_range(ref))
    validation = _build_validation(rule, sheet_names(sheet))
    validation.add(target)
    sheet.add_data_validation(validation)
    return target


def _build_validation(rule: DataValidationRule, names: list[str]) -> DataValidation:
    messages = {
        "allow_blank": rule.allow_blank,
        "error": rule.error_message,
        "showErrorMessage": True,
        "prompt": rule.prompt,
        "showInputMessage": rule.prompt is not None,
    }
    match rule.type:
        case "list":
            if not rule.options:
                raise InvalidArgumentError("list validation needs options.")
            if any("," in option or '"' in option for option in rule.options):
                raise InvalidArgumentError("list options cannot contain commas or quotes.")
            joined = ",".join(rule.options)
            if len(joined) > 255:
                raise InvalidArgumentError("list options can be at most 255 characters in total.")
            return DataValidation(type="list", formula1=f'"{joined}"', **messages)
        case "custom":
            if not rule.formula:
                raise InvalidArgumentError("custom validation needs a formula.")
            formula = storable_operand(rule.formula, names)
            return DataValidation(type="custom", formula1=formula, **messages)
        case _:
            if rule.operator is None or rule.minimum is None:
                raise InvalidArgumentError(f"{rule.type} validation needs an operator and minimum.")
            return DataValidation(
                type="textLength" if rule.type == "text_length" else rule.type,
                operator=rule.operator,
                formula1=_checked_operand(rule.minimum, names),
                formula2=_checked_operand(rule.maximum, names) if rule.maximum else None,
                **messages,
            )


def _checked_operand(value: str, names: list[str]) -> str:
    """Rule operands are stored as formulas without the leading '='."""
    return storable_operand(value, names)
