"""Calculated fields: a formula over source fields, applied to their sums as Excel does."""

from collections.abc import Callable
from dataclasses import dataclass

from openpyxl import Workbook
from openpyxl.formula import Tokenizer
from openpyxl.formula.tokenizer import Token, TokenizerError

from excel_mcp.calc.engine import Engine
from excel_mcp.calc.values import ExcelError, UncalculableError
from excel_mcp.errors import InvalidArgumentError
from excel_mcp.formulas import check_formula, unquote
from excel_mcp.operations.pivot_options import CalculatedField
from excel_mcp.operations.pivot_source import Source

FieldKey = tuple[str, int]
"""("column", index) for a source column, ("calculated", index) for an earlier calculated field."""
Number = float | int | ExcelError | None


@dataclass(frozen=True)
class CalcField:
    name: str
    formula: str
    parts: list[str | FieldKey]


def compile_calculated(source: Source, fields: list[CalculatedField]) -> list[CalcField]:
    taken = {column.name.strip().casefold() for column in source.columns}
    compiled: list[CalcField] = []
    for field in fields:
        name = field.name.strip()
        if not name or name.casefold() in taken:
            raise InvalidArgumentError(
                f"Calculated field name {field.name!r} is empty or already used by a field."
            )
        taken.add(name.casefold())
        formula = field.formula.strip().removeprefix("=")
        compiled.append(CalcField(name, formula, _parts(source, compiled, formula)))
    return compiled


def _parts(source: Source, earlier: list[CalcField], formula: str) -> list[str | FieldKey]:
    try:
        tokens = Tokenizer(f"={formula}").items
    except TokenizerError as error:
        raise InvalidArgumentError(f"Formula {formula!r} could not be parsed: {error}.") from None
    parts: list[str | FieldKey] = []
    for token in tokens:
        if token.type == Token.OPERAND and token.subtype == Token.RANGE:
            parts.append(_field(source, earlier, token.value, formula))
        else:
            parts.append(token.value)
    # The field names are checked above; the rest must pass the safety check as written.
    check_formula("=" + "".join("1" if isinstance(part, tuple) else part for part in parts), [])
    return parts


def _field(source: Source, earlier: list[CalcField], text: str, formula: str) -> FieldKey:
    name = unquote(text).strip().casefold()
    for index, column in enumerate(source.columns):
        if column.name.strip().casefold() == name:
            if column.kind != "number":
                raise InvalidArgumentError(
                    f"Field {column.name!r} holds {column.kind}, not numbers, so a calculated "
                    "field cannot use it."
                )
            return "column", index
    for index, calc in enumerate(earlier):
        if calc.name.casefold() == name:
            return "calculated", index
    available = [column.name for column in source.columns] + [calc.name for calc in earlier]
    raise InvalidArgumentError(
        f"{text!r} in the formula {formula!r} is not a field. Fields: {available}. "
        "Quote names that contain spaces, like 'Unit Price'."
    )


def references(calc: CalcField) -> list[FieldKey]:
    return [part for part in calc.parts if isinstance(part, tuple)]


class Evaluator:
    """Calculates formulas with the engine that reads formula results."""

    def __init__(self) -> None:
        self.workbook = Workbook()
        self.empty = Workbook()

    def evaluate(self, calc: CalcField, value_of: Callable[[FieldKey], Number]) -> Number:
        numbers: dict[FieldKey, float] = {}
        for key in references(calc):
            value = value_of(key)
            if isinstance(value, ExcelError):
                return value
            numbers[key] = float(value or 0)
        text = "".join(
            part if isinstance(part, str) else f"({numbers[part]!r})" for part in calc.parts
        )
        sheet = self.workbook.worksheets[0]
        sheet["A1"] = f"={text}"
        try:
            result = Engine(self.empty, self.workbook).calculate(sheet, 1, 1)
        except UncalculableError as reason:
            raise InvalidArgumentError(
                f"Cannot calculate the field {calc.name!r}: {reason}. Use plain arithmetic "
                "and functions such as IF."
            ) from None
        if isinstance(result, ExcelError | float | int):
            return result
        raise InvalidArgumentError(f"The field {calc.name!r} must calculate a number.")
