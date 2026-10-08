"""Data validation rules."""

import datetime as dt
import re
from typing import Literal

from openpyxl.utils.datetime import to_excel
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.worksheet import Worksheet
from pydantic import Field

from excel_mcp.errors import InvalidArgumentError
from excel_mcp.formulas import storable_operand
from excel_mcp.inputs import InputModel
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

_ISO_DATE = re.compile(r"\d{4}-\d{2}-\d{2}")
_CLOCK_TIME = re.compile(r"\d{1,2}:\d{2}(:\d{2})?")
_AREA = re.compile(r"\$?[A-Za-z]{1,3}\$?\d+:\$?[A-Za-z]{1,3}\$?\d+")


class DataValidationRule(InputModel):
    """Which fields apply depends on ``type``."""

    type: Literal["list", "whole", "decimal", "date", "time", "text_length", "custom"]
    options: list[str] | None = Field(default=None, description="list: allowed values, typed in.")
    source: str | None = Field(
        default=None,
        description="list: cells or a name holding the allowed values, instead of options, "
        "e.g. '=$A$2:$A$20', '=Sheet2!$A:$A' or '=Regions'.",
    )
    operator: Operator | None = Field(
        default=None, description="whole, decimal, date, time, text_length."
    )
    minimum: str | None = Field(
        default=None,
        description="Number or formula; dates as '2026-01-31', times as '09:30'.",
    )
    maximum: str | None = Field(default=None, description="For between/notBetween.")
    formula: str | None = Field(
        default=None, description="custom: for the range's top-left cell, e.g. '=A2>B2'."
    )
    allow_blank: bool = True
    prompt_title: str | None = Field(
        default=None, max_length=32, description="Input message title."
    )
    prompt: str | None = Field(
        default=None, max_length=255, description="Input message, shown when a cell is selected."
    )
    error_style: Literal["stop", "warning", "information"] = Field(
        default="stop", description="stop rejects bad input; the others let the user keep it."
    )
    error_title: str | None = Field(default=None, max_length=32)
    error_message: str | None = Field(default=None, max_length=255)


def add_data_validation(sheet: Worksheet, ref: str, rule: DataValidationRule) -> str:
    target = str(parse_range(ref))
    validation = _build_validation(rule, sheet_names(sheet))
    validation.add(target)
    sheet.add_data_validation(validation)
    return target


def _build_validation(rule: DataValidationRule, names: list[str]) -> DataValidation:
    messages = {
        "allow_blank": rule.allow_blank,
        "errorStyle": rule.error_style,
        "errorTitle": rule.error_title,
        "error": rule.error_message,
        "showErrorMessage": True,
        "promptTitle": rule.prompt_title,
        "prompt": rule.prompt,
        "showInputMessage": bool(rule.prompt_title or rule.prompt),
    }
    if rule.type != "list" and (rule.options or rule.source):
        raise InvalidArgumentError("options and source are only for list validation.")
    match rule.type:
        case "list":
            return DataValidation(type="list", formula1=_list_source(rule, names), **messages)
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
                formula1=_operand(rule.minimum, rule.type, names),
                formula2=_operand(rule.maximum, rule.type, names) if rule.maximum else None,
                **messages,
            )


def _list_source(rule: DataValidationRule, names: list[str]) -> str:
    if bool(rule.options) == bool(rule.source):
        raise InvalidArgumentError("list validation needs either options or source.")
    if rule.source:
        source = storable_operand(rule.source, names)
        area = source.rpartition("!")[2]
        if _AREA.fullmatch(area) and min(parse_range(area).rows, parse_range(area).cols) > 1:
            raise InvalidArgumentError(f"A list source must be one row or column, not {area}.")
        return source
    options = rule.options or []
    if any("," in option or '"' in option for option in options):
        raise InvalidArgumentError("list options cannot contain commas or quotes.")
    joined = ",".join(options)
    if len(joined) > 255:
        raise InvalidArgumentError("list options can be at most 255 characters in total.")
    return f'"{joined}"'


def _operand(value: str, kind: str, names: list[str]) -> str:
    """Operands are stored as formulas without the '='; dates and times as serial numbers."""
    try:
        if kind == "date" and _ISO_DATE.fullmatch(value):
            return str(int(to_excel(dt.date.fromisoformat(value))))
        if kind == "time" and _CLOCK_TIME.fullmatch(value):
            padded = value if value.count(":") == 2 else value + ":00"
            return repr(to_excel(dt.time.fromisoformat(padded.zfill(8))))
    except ValueError:
        raise InvalidArgumentError(f"{value!r} is not a valid {kind}.") from None
    return storable_operand(value, names)
