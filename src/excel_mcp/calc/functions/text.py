"""Text functions."""

import re

from excel_mcp.calc.functions.helpers import flat
from excel_mcp.calc.registry import function
from excel_mcp.calc.textformat import format_fixed, format_number
from excel_mcp.calc.values import (
    MAX_TEXT,
    VALUE,
    FormulaError,
    Scalar,
    Value,
    parse_number,
    to_bool,
    to_int,
    to_number,
    to_text,
)
from excel_mcp.calc.wildcard import Wildcard


def _count(value: Scalar) -> int:
    count = to_int(value)
    if count < 0:
        raise FormulaError(VALUE)
    return count


@function("HYPERLINK", kind="scalar")
def hyperlink(link: Scalar, friendly_name: Scalar = ...) -> Scalar:  # pyright: ignore[reportArgumentType]
    return to_text(link) if friendly_name is ... else friendly_name


@function("LEN", kind="scalar")
def len_(text: Scalar) -> float:
    return float(len(to_text(text)))


@function("LEFT", kind="scalar")
def left(text: Scalar, count: Scalar = 1) -> str:
    return to_text(text)[: _count(count)]


@function("RIGHT", kind="scalar")
def right(text: Scalar, count: Scalar = 1) -> str:
    n = _count(count)
    return to_text(text)[-n:] if n else ""


@function("MID", kind="scalar")
def mid(text: Scalar, start: Scalar, count: Scalar) -> str:
    first = to_int(start)
    if first < 1:
        raise FormulaError(VALUE)
    return to_text(text)[first - 1 : first - 1 + _count(count)]


@function("UPPER", kind="scalar")
def upper(text: Scalar) -> str:
    return to_text(text).upper()


@function("LOWER", kind="scalar")
def lower(text: Scalar) -> str:
    return to_text(text).lower()


@function("PROPER", kind="scalar")
def proper(text: Scalar) -> str:
    return re.sub(r"[^\W\d_]+", lambda m: m.group().capitalize(), to_text(text))


@function("TRIM", kind="scalar")
def trim(text: Scalar) -> str:
    return re.sub(" +", " ", to_text(text).strip(" "))


@function("EXACT", kind="scalar")
def exact(first: Scalar, second: Scalar) -> bool:
    return to_text(first) == to_text(second)


@function("CONCAT")
def concat(*args: Value) -> str:
    return _limited("".join(to_text(v) for v in flat(args)))


@function("CONCATENATE", kind="scalar")
def concatenate(*args: Scalar) -> str:
    return _limited("".join(to_text(v) for v in args))


@function("TEXTJOIN")
def textjoin(delimiter: Value, ignore_empty: Value, *args: Value) -> str:
    separator = to_text(delimiter)
    skip = to_bool(ignore_empty)
    items = [to_text(v) for v in flat(args)]
    return _limited(separator.join(i for i in items if i or not skip))


def _limited(text: str) -> str:
    if len(text) > MAX_TEXT:
        raise FormulaError(VALUE)
    return text


@function("SUBSTITUTE", kind="scalar")
def substitute(text: Scalar, old: Scalar, new: Scalar, instance: Scalar = None) -> str:
    source, target, replacement = to_text(text), to_text(old), to_text(new)
    if not target:
        return source
    if instance is None:
        grown = source.count(target) * (len(replacement) - len(target))
        if len(source) + grown > MAX_TEXT:
            raise FormulaError(VALUE)
        return source.replace(target, replacement)
    which = to_int(instance)
    if which < 1:
        raise FormulaError(VALUE)
    position = -1
    for _ in range(which):
        position = source.find(target, position + 1)
        if position < 0:
            return source
    return source[:position] + replacement + source[position + len(target) :]


def _start(within: str, start: Scalar) -> int:
    first = to_int(start)
    if first < 1 or first > len(within) + (0 if within else 1):
        raise FormulaError(VALUE)
    return first - 1


@function("FIND", kind="scalar")
def find(needle: Scalar, text: Scalar, start: Scalar = 1) -> float:
    within = to_text(text)
    index = within.find(to_text(needle), _start(within, start))
    if index < 0:
        raise FormulaError(VALUE)
    return float(index + 1)


@function("SEARCH", kind="scalar")
def search(needle: Scalar, text: Scalar, start: Scalar = 1) -> float:
    within = to_text(text)
    pattern = to_text(needle)
    begin = _start(within, start)
    index = Wildcard(pattern).search(within, begin)
    if index is None:
        raise FormulaError(VALUE)
    return float(index + 1)


@function("REPT", kind="scalar")
def rept(text: Scalar, count: Scalar) -> str:
    source, times = to_text(text), _count(count)
    if len(source) * times > MAX_TEXT:
        raise FormulaError(VALUE)
    return source * times


@function("VALUE", kind="scalar")
def value(text: Scalar) -> float:
    if isinstance(text, str):
        number = parse_number(text)
        if number is None:
            raise FormulaError(VALUE)
        return number
    return to_number(text)


@function("NUMBERVALUE", kind="scalar")
def numbervalue(text: Scalar, decimal: Scalar = ".", group: Scalar = ",") -> float:
    body = re.sub(r"\s", "", to_text(text))
    decimal_mark, group_mark = to_text(decimal), to_text(group)
    if len(decimal_mark) != 1 or len(group_mark) != 1 or decimal_mark == group_mark:
        raise FormulaError(VALUE)
    percents = len(body) - len(body.rstrip("%"))
    body = body.rstrip("%")
    whole, _, fraction = body.partition(decimal_mark)
    if decimal_mark in fraction:
        raise FormulaError(VALUE)
    cleaned = whole.replace(group_mark, "") + ("." + fraction if decimal_mark in body else "")
    if not cleaned:
        return 0.0
    if not re.fullmatch(r"[+-]?(\d+\.?\d*|\.\d+)", cleaned):
        raise FormulaError(VALUE)
    return float(cleaned) / 100**percents


@function("TEXT", kind="scalar")
def text_(number: Scalar, fmt: Scalar) -> str:
    if isinstance(number, bool):
        return to_text(number)
    if to_text(fmt) == "":
        return ""
    if isinstance(number, str):
        parsed = parse_number(number)
        if parsed is None:
            return number
        number = parsed
    return format_number(to_number(number), to_text(fmt))


@function("DOLLAR", kind="scalar")
def dollar(number: Scalar, decimals: Scalar = 2) -> str:
    n = to_number(number)
    shown = format_fixed(abs(n), to_int(decimals), grouping=True)
    return f"(${shown})" if n < 0 else f"${shown}"


@function("FIXED", kind="scalar")
def fixed(number: Scalar, decimals: Scalar = 2, no_commas: Scalar = False) -> str:
    return format_fixed(to_number(number), to_int(decimals), grouping=not to_bool(no_commas))


@function("CHAR", kind="scalar")
def char(number: Scalar) -> str:
    code = to_int(number)
    if not 1 <= code <= 255:
        raise FormulaError(VALUE)
    return bytes([code]).decode("cp1252", errors="strict")


@function("CODE", kind="scalar")
def code(text: Scalar) -> float:
    first = to_text(text)[:1]
    if not first:
        raise FormulaError(VALUE)
    return float(first.encode("cp1252")[0])


@function("UNICHAR", kind="scalar")
def unichar(number: Scalar) -> str:
    code = to_int(number)
    if not 1 <= code <= 0x10FFFF or 0xD800 <= code <= 0xDFFF:
        raise FormulaError(VALUE)
    return chr(code)


@function("UNICODE", kind="scalar")
def unicode_(text: Scalar) -> float:
    first = to_text(text)[:1]
    if not first:
        raise FormulaError(VALUE)
    return float(ord(first))


@function("REPLACE", kind="scalar")
def replace(text: Scalar, start: Scalar, count: Scalar, new: Scalar) -> str:
    source, first = to_text(text), to_int(start)
    if first < 1 or to_int(count) < 0:
        raise FormulaError(VALUE)
    return _limited(source[: first - 1] + to_text(new) + source[first - 1 + to_int(count) :])


@function("CLEAN", kind="scalar")
def clean(text: Scalar) -> str:
    return "".join(c for c in to_text(text) if ord(c) >= 32)
