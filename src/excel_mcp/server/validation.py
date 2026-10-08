"""One readable line per problem in place of pydantic's validation report.

Input values are never echoed in full: they may be large or hold cell data.
"""

import difflib
from collections.abc import Iterator

from pydantic import ValidationError
from pydantic_core import ErrorDetails

from excel_mcp.text import quoted

MAX_PROBLEMS = 10
MAX_VALUE_LENGTH = 30

_TYPE_NAMES = {
    "string_type": "text",
    "int_type": "number",
    "int_parsing": "number",
    "float_type": "number",
    "float_parsing": "number",
    "bool_type": "boolean",
    "bool_parsing": "boolean",
    "none_required": "null",
    "list_type": "list",
    "dict_type": "object",
    "model_type": "object",
    "model_attributes_type": "object",
}
_MEMBER_TAGS = {"str", "int", "float", "bool", "none"}
_LIMITS = {
    "greater_than_equal": ("at least", "ge"),
    "greater_than": ("greater than", "gt"),
    "less_than_equal": ("at most", "le"),
    "less_than": ("less than", "lt"),
}
_LENGTHS = {
    "too_short": ("at least", "min_length"),
    "too_long": ("at most", "max_length"),
    "string_too_short": ("at least", "min_length"),
    "string_too_long": ("at most", "max_length"),
}


def format_validation_error(tool: str, error: ValidationError) -> str:
    problems = list(dict.fromkeys(_problems(error.errors())))
    shown = problems[:MAX_PROBLEMS]
    if len(problems) > MAX_PROBLEMS:
        shown.append(f"and {len(problems) - MAX_PROBLEMS} more.")
    header = f"Invalid arguments for {tool}:"
    if len(shown) == 1:
        return f"{header} {shown[0]}"
    return "\n".join([header, *(f"- {problem}" for problem in shown)])


def _problems(errors: list[ErrorDetails]) -> Iterator[str]:
    # A value that matches no member of a union fails once per member, at "<path>.<member>".
    members: dict[tuple[int | str, ...], list[ErrorDetails]] = {}
    for detail in errors:
        if _is_union_member(detail):
            members.setdefault(detail["loc"][:-1], []).append(detail)
        elif detail["type"] == "unknown_fields":
            yield from _unknown_fields(detail)
        else:
            yield _line(detail["loc"], _message(detail))
    for loc, details in members.items():
        names = dict.fromkeys(_TYPE_NAMES[detail["type"]] for detail in details)
        yield _line(loc, f"expected {_either(list(names))}; got {_kind(details[0]['input'])}")


def _is_union_member(detail: ErrorDetails) -> bool:
    return detail["type"] in _TYPE_NAMES and str(detail["loc"][-1]) in _MEMBER_TAGS


def _unknown_fields(detail: ErrorDetails) -> Iterator[str]:
    context = detail.get("ctx", {})
    for name in context["unknown"]:
        close = difflib.get_close_matches(name, context["valid"], n=1)
        hint = (
            f"did you mean {close[0]!r}?" if close else f"valid fields: {quoted(context['valid'])}"
        )
        yield _line((*detail["loc"], name), f"unknown field; {hint}")


def _message(detail: ErrorDetails) -> str:
    kind, context = detail["type"], detail.get("ctx", {})
    if kind == "missing":
        return "required"
    if kind == "literal_error":
        return f"{_short(detail['input'])} is not valid; use {context['expected']}."
    if kind in _LIMITS:
        wording, key = _LIMITS[kind]
        return f"must be {wording} {context[key]}"
    if kind in _LENGTHS:
        wording, key = _LENGTHS[kind]
        unit = "characters" if kind.startswith("string") else "items"
        return f"must have {wording} {context[key]} {unit}"
    if kind in _TYPE_NAMES:
        return f"expected {_TYPE_NAMES[kind]}; got {_kind(detail['input'])}"
    if kind in ("value_error", "assertion_error"):
        return detail["msg"].removeprefix("Value error, ").removeprefix("Assertion failed, ")
    return f"{detail['msg'][:1].lower()}{detail['msg'][1:]} (got {_short(detail['input'])})"


def _line(loc: tuple[int | str, ...], message: str) -> str:
    path = ""
    for part in loc:
        path += f"[{part}]" if isinstance(part, int) else f"{'.' if path else ''}{part}"
    return f"{path}: {message}" if path else message


def _either(names: list[str]) -> str:
    return names[0] if len(names) == 1 else f"{', '.join(names[:-1])} or {names[-1]}"


def _kind(value: object) -> str:
    if value is None:
        return "null"
    if isinstance(value, bool):
        return "boolean"
    if isinstance(value, int | float):
        return "number"
    if isinstance(value, str):
        return "text"
    return "list" if isinstance(value, list) else "object"


def _short(value: object) -> str:
    text = repr(value)
    return text if len(text) <= MAX_VALUE_LENGTH else f"{text[: MAX_VALUE_LENGTH - 1]}…"
