"""Syntax check for formulas, so that text Excel cannot parse is never written to a file.

Excel's parser forgives a person typing (it closes a missing ")" for them), but a file
that holds such a formula is reported as damaged. The checks follow the grammar:
balanced brackets, operators with operands on both sides, arrays of constants, and
references or names that can exist. A function Excel knows must get a number of arguments
it accepts (`function_arguments.json`, made by `scripts/probe_function_arguments.py`);
other names, such as user-defined functions, are not checked.
"""

import json
import re
from dataclasses import dataclass, field
from pathlib import Path

from openpyxl.formula.tokenizer import Token

from excel_mcp.errors import InvalidFormulaError

_NAME = r"(?:[^\W\d]|\\)[\w.?\\]*"
_CELL = r"\$?[A-Za-z]{1,3}\$?\d+"
_PART = rf"(?:{_CELL}|\$?[A-Za-z]{{1,3}}|\$?\d+|{_NAME})"
_SHEET = r"(?:'(?:[^']|'')+'|[^\s'!\[\]:*?/\\]+)"
_QUALIFIER = rf"(?:(?:\[\d+\])?{_SHEET}(?::{_SHEET})?!)?"
_QUALIFIED_PART = rf"{_QUALIFIER}{_PART}"
_REFERENCE = re.compile(
    rf"@*(?:{_QUALIFIER}(?:{_CELL}|{_NAME}|#REF!)|{_QUALIFIED_PART}(?::{_QUALIFIED_PART})+)"
)
_TABLE_REFERENCE = re.compile(rf"@*(?:(?:{_SHEET}!)?{_NAME})?(\[.*\])")
_ARGUMENTS: dict[str, tuple[int, int, int]] = {
    name: (limits[0], limits[1], limits[2])
    for name, limits in json.loads(
        Path(__file__).with_name("function_arguments.json").read_text(encoding="utf-8")
    ).items()
}
_FUNCTION_PREFIXES = ("_XLFN.", "_XLWS.")
_ARRAY_VALUES = {Token.NUMBER, Token.TEXT, Token.LOGICAL, Token.ERROR}
_TABLE_ITEMS = {"#all", "#data", "#headers", "#totals", "#this row"}

_ARRAY_ONLY = "array constants can only hold numbers, text, TRUE, FALSE and errors."
_KINDS = {Token.FUNC: "function", Token.PAREN: "paren", Token.ARRAY: "array"}

# What the previous token leaves the next one to be.
_VALUE, _OPERATOR, _EMPTY_OK, _NEEDS_VALUE = "value", "operator", "empty_ok", "needs_value"


class _ProblemError(Exception):
    pass


@dataclass
class _Open:
    """A bracket that is not closed yet; arrays also count the elements of each row."""

    kind: str
    function: str = ""
    separators: int = 0
    row_lengths: list[int] = field(default_factory=lambda: [1])


def check_syntax(formula: str, tokens: list[Token]) -> None:
    """Raise `InvalidFormulaError` unless the tokens of ``formula`` form a valid formula."""
    try:
        _Checker(tokens).run()
    except _ProblemError as problem:
        raise invalid_formula(formula, str(problem).removesuffix(".")) from None


def invalid_formula(formula: str, reason: str) -> InvalidFormulaError:
    shown = formula if len(formula) <= 60 else f"{formula[:59]}…"
    return InvalidFormulaError(f"Formula {shown!r} is not valid: {reason}.")


class _Checker:
    def __init__(self, tokens: list[Token]) -> None:
        self.tokens = tokens
        self.stack: list[_Open] = []
        self.previous = _EMPTY_OK
        self.previous_text = ""
        self.reference_ended = False
        self.spaced = False
        self.invalid_operand: _ProblemError | None = None

    def run(self) -> None:
        if all(token.type == Token.WSPACE for token in self.tokens):
            raise _ProblemError("it is empty.")
        for token in self.tokens:
            if token.type == Token.WSPACE:
                self.spaced = True
                continue
            self._step(token)
            self.spaced = False
        if self.stack:
            bracket = "{" if self.stack[-1].kind == "array" else "("
            raise _ProblemError(f"unclosed {bracket!r}.")
        if self.previous == _OPERATOR:
            raise _ProblemError(f"{self.previous_text!r} needs a value after it.")
        if self.invalid_operand:
            raise self.invalid_operand

    def _step(self, token: Token) -> None:
        kind, subtype, value = token.type, token.subtype, token.value
        if kind == Token.OPERAND:
            self._operand(token)
        elif kind == Token.FUNC and subtype == Token.OPEN:
            _check_function_name(value)
            if not self._continues_range(value):
                self._begin_value(value)
            self._open("function", _EMPTY_OK, _function_name(value))
        elif kind == Token.PAREN and subtype == Token.OPEN:
            self._begin_value(value)
            self._open("paren", _NEEDS_VALUE)
        elif kind == Token.ARRAY and subtype == Token.OPEN:
            self._begin_value(value)
            self._open("array", _NEEDS_VALUE)
        elif kind in (Token.FUNC, Token.PAREN, Token.ARRAY):
            self._close(kind, value)
        elif kind == Token.SEP:
            self._separator(subtype, value)
        elif kind == Token.OP_PRE:
            self._prefix(value)
        elif kind == Token.OP_IN:
            self._infix(value)
        elif kind == Token.OP_POST and self.previous != _VALUE:
            raise _ProblemError(f"{value!r} needs a value before it.")

    def _in_array(self) -> bool:
        return bool(self.stack) and self.stack[-1].kind == "array"

    def _begin_value(self, text: str) -> None:
        """Check that a value may start here; two values in a row need an intersection."""
        if self._in_array():
            raise _ProblemError(_ARRAY_ONLY)
        if self.previous == _VALUE and not (self.spaced and self.reference_ended):
            raise _ProblemError(f"an operator is missing before {text!r}.")

    def _operand(self, token: Token) -> None:
        if token.subtype == Token.RANGE and set(token.value) == {"@"}:
            self._prefix("@")
            return
        if self._in_array() and token.subtype not in _ARRAY_VALUES:
            raise _ProblemError(_ARRAY_ONLY)
        continues_range = self._continues_range(token.value)
        text = token.value.removeprefix(":") if continues_range else token.value
        if (
            not continues_range
            and self.previous == _VALUE
            and not (self.spaced and self.reference_ended and _is_reference_token(token))
        ):
            raise _ProblemError(f"an operator is missing before {token.value!r}.")
        if token.subtype == Token.RANGE and not _is_reference(text):
            self.invalid_operand = self.invalid_operand or _ProblemError(
                f"{token.value!r} is not a valid reference, name or number."
            )
        self._ends_value(reference=_is_reference_token(token))

    def _continues_range(self, text: str) -> bool:
        """Whether ``text`` carries on a reference that just ended, as in INDEX(...):A3."""
        return (
            text.startswith(":")
            and self.previous == _VALUE
            and self.reference_ended
            and not self.spaced
        )

    def _ends_value(self, reference: bool) -> None:
        self.previous = _VALUE
        self.reference_ended = reference

    def _open(self, kind: str, next_previous: str, function: str = "") -> None:
        self.stack.append(_Open(kind, function))
        self.previous = next_previous

    def _close(self, token_type: str, value: str) -> None:
        if not self.stack or self.stack[-1].kind != _KINDS[token_type]:
            raise _ProblemError(f"unmatched {value!r}.")
        top = self.stack[-1]
        self._needs_no_more()
        if self.previous == _NEEDS_VALUE:
            what = "an array element" if top.kind == "array" else "an expression"
            raise _ProblemError(f"{what} is missing before {value!r}.")
        if top.kind == "function":
            self._check_argument_count(top)
        if len(set(top.row_lengths)) > 1:
            raise _ProblemError("every row of an array constant needs the same number of elements.")
        self.stack.pop()
        self._ends_value(reference=top.kind != "array")

    def _check_argument_count(self, call: _Open) -> None:
        if call.function not in _ARGUMENTS:
            return
        count = 0 if call.separators == 0 and self.previous == _EMPTY_OK else call.separators + 1
        fewest, most, step = _ARGUMENTS[call.function]
        if count < fewest:
            raise _ProblemError(
                f"{call.function} needs at least {_arguments(fewest)}, got {count}."
            )
        if count > most:
            raise _ProblemError(f"{call.function} takes at most {_arguments(most)}, got {count}.")
        if (count - fewest) % step:
            raise _ProblemError(
                f"{call.function} takes {fewest}, {fewest + step}, {fewest + 2 * step}... "
                f"arguments, got {count}."
            )

    def _separator(self, subtype: str, value: str) -> None:
        top = self.stack[-1] if self.stack else None
        in_array = top is not None and top.kind == "array"
        if top is None or top.kind == "paren" or (subtype == Token.ROW and not in_array):
            raise _ProblemError(f"unexpected {value!r}.")
        self._needs_no_more()
        if not in_array:
            top.separators += 1
            self.previous = _EMPTY_OK
            return
        if self.previous != _VALUE:
            raise _ProblemError("array constants cannot have empty elements.")
        if subtype == Token.ROW:
            top.row_lengths.append(1)
        else:
            top.row_lengths[-1] += 1
        self.previous = _NEEDS_VALUE

    def _needs_no_more(self) -> None:
        if self.previous == _OPERATOR:
            raise _ProblemError(f"{self.previous_text!r} needs a value after it.")

    def _prefix(self, value: str) -> None:
        if self.previous == _VALUE:
            raise _ProblemError(f"an operator is missing before {value!r}.")
        if self._in_array() and value not in ("+", "-"):
            raise _ProblemError(_ARRAY_ONLY)
        self.previous, self.previous_text = _OPERATOR, value

    def _infix(self, value: str) -> None:
        if self._in_array():
            raise _ProblemError(_ARRAY_ONLY)
        if self.previous != _VALUE:
            raise _ProblemError(f"{value!r} needs a value before it.")
        if value == "," and not self.reference_ended:
            raise _ProblemError("',' can only join references or separate function arguments.")
        self.previous, self.previous_text = _OPERATOR, value


def _arguments(count: int) -> str:
    return f"{count} argument{'' if count == 1 else 's'}"


def _is_reference_token(token: Token) -> bool:
    """A reference, or the #REF! that Excel leaves where one was deleted."""
    return token.subtype == Token.RANGE or token.value == "#REF!"


def _function_name(token_value: str) -> str:
    """The bare, upper-case name of a function token such as ``A1:@_xlfn.IFS(``."""
    name = token_value.removesuffix("(").rpartition(":")[2].rpartition("!")[2].lstrip("@").upper()
    while name.startswith(_FUNCTION_PREFIXES):
        name = name.split(".", 1)[1]
    return name


def _check_function_name(token_value: str) -> None:
    name = token_value.removesuffix("(").rpartition(":")[2].rpartition("!")[2].lstrip("@")
    if name and not re.fullmatch(_NAME, name):
        raise _ProblemError(f"{token_value!r} is not a valid function name.")


def _is_reference(text: str) -> bool:
    if "[" in text:
        return _is_table_reference(text)
    return _REFERENCE.fullmatch(text) is not None


def _is_table_reference(text: str) -> bool:
    if "[[]" in text or not (_TABLE_REFERENCE.fullmatch(text) or re.fullmatch(r"\[\d+\].*", text)):
        return False
    depth = 0
    escaped = False
    for character in text:
        if escaped:
            escaped = False
        elif character == "'":
            escaped = True
        elif character == "[":
            depth += 1
        elif character == "]":
            depth -= 1
            if depth < 0:
                return False
    items = re.findall(r"\[(#[^\]]*)\]", text)
    return depth == 0 and all(item.casefold() in _TABLE_ITEMS for item in items)
