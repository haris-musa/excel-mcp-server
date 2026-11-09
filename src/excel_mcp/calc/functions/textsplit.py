"""TEXTBEFORE, TEXTAFTER, TEXTSPLIT, VALUETOTEXT and ARRAYTOTEXT."""

from excel_mcp.calc.functions.helpers import as_grid
from excel_mcp.calc.registry import function
from excel_mcp.calc.values import (
    CALC,
    NA,
    VALUE,
    ExcelError,
    FormulaError,
    Grid,
    Scalar,
    Value,
    check_array_size,
    scalar,
    to_bool,
    to_int,
    to_text,
)


def _delimiters(value: Value) -> list[str]:
    items = as_grid(value).flat()
    for item in items:
        if isinstance(item, ExcelError):
            raise FormulaError(item)
    return [to_text(item) for item in items]  # pyright: ignore[reportArgumentType]


def _positions(text: str, delimiters: list[str], ignore_case: bool) -> list[tuple[int, int]]:
    """Where delimiters start and end in the text, left to right, never overlapping."""
    haystack = text.casefold() if ignore_case else text
    needles = sorted(
        {d.casefold() if ignore_case else d for d in delimiters if d}, key=len, reverse=True
    )
    found, index = [], 0
    while index <= len(haystack):
        match = next((n for n in needles if haystack.startswith(n, index)), None)
        if match is None:
            index += 1
        else:
            found.append((index, index + len(match)))
            index += len(match)
    return found


def _cut(
    text: Value,
    delimiter: Value,
    instance: Value,
    match_mode: Value,
    match_end: Value,
    if_not_found: Value,
    *,
    before: bool,
) -> Scalar:
    source = to_text(scalar(text))
    delimiters = _delimiters(delimiter)
    count = 1 if instance is None else to_int(scalar(instance))
    ignore_case = False if match_mode is None else to_int(scalar(match_mode)) == 1
    at_end = False if match_end is None else to_bool(scalar(match_end))
    shortest = min((len(d) for d in delimiters if d), default=0)
    if count == 0 or (
        shortest and instance is not None and abs(count) > len(source) - shortest + 1
    ):
        raise FormulaError(VALUE)
    if match_mode is not None and to_int(scalar(match_mode)) not in (0, 1):
        raise FormulaError(VALUE)
    if all(d == "" for d in delimiters):
        position = len(source) if count < 0 or len(source) == 1 else 0
        return source[:position] if before else source[position:]
    found = _positions(source, delimiters, ignore_case)
    if at_end:
        found = [*found, (len(source), len(source))] if count > 0 else [(0, 0), *found]
    if len(found) < abs(count):
        if if_not_found is None or if_not_found is ...:
            raise FormulaError(NA)
        return scalar(if_not_found)
    start, end = found[count - 1] if count > 0 else found[count]
    return source[:start] if before else source[end:]


@function("TEXTBEFORE", kind="raw", array=(1,))
def textbefore(
    text: Value,
    delimiter: Value,
    instance: Value = None,
    match_mode: Value = None,
    match_end: Value = None,
    if_not_found: Value = ...,  # pyright: ignore[reportArgumentType]
) -> Scalar:
    _check(text, delimiter)
    return _cut(text, delimiter, instance, match_mode, match_end, if_not_found, before=True)


@function("TEXTAFTER", kind="raw", array=(1,))
def textafter(
    text: Value,
    delimiter: Value,
    instance: Value = None,
    match_mode: Value = None,
    match_end: Value = None,
    if_not_found: Value = ...,  # pyright: ignore[reportArgumentType]
) -> Scalar:
    _check(text, delimiter)
    return _cut(text, delimiter, instance, match_mode, match_end, if_not_found, before=False)


def _check(*values: Value) -> None:
    for value in values:
        if isinstance(value, ExcelError):
            raise FormulaError(value)


def _split(text: str, delimiters: list[str], ignore_case: bool, ignore_empty: bool) -> list[str]:
    pieces, last = [], 0
    for start, end in _positions(text, delimiters, ignore_case):
        pieces.append(text[last:start])
        last = end
    pieces.append(text[last:])
    return [p for p in pieces if p] if ignore_empty else pieces


@function("TEXTSPLIT", kind="raw", array=(1, 2))
def textsplit(
    text: Value,
    column_delimiter: Value,
    row_delimiter: Value = None,
    ignore_empty: Value = False,
    match_mode: Value = 0,
    pad_with: Value = None,
) -> Grid:
    _check(text, column_delimiter, row_delimiter)
    source = to_text(scalar(text))
    if source == "":
        raise FormulaError(VALUE)
    mode = to_int(scalar(match_mode))
    if mode not in (0, 1):
        raise FormulaError(VALUE)
    skip = to_bool(scalar(ignore_empty))
    columns = _delimiters(column_delimiter) if column_delimiter is not None else []
    rows = _delimiters(row_delimiter) if row_delimiter is not None else []
    if not [d for d in columns + rows if d] or "" in columns + rows:
        raise FormulaError(VALUE)
    lines = _split(source, rows, mode == 1, skip) if any(rows) else [source]
    cells = [_split(line, columns, mode == 1, skip) if any(columns) else [line] for line in lines]
    width = max((len(line) for line in cells), default=0)
    if width == 0:
        raise FormulaError(CALC)
    check_array_size(len(cells), width)
    filler: Scalar = NA if pad_with is None else scalar(pad_with)
    return Grid([list(line) + [filler] * (width - len(line)) for line in cells])


def _shown(value: Scalar, strict: bool) -> str:
    if isinstance(value, ExcelError):
        return value.code
    if isinstance(value, str):
        return '"' + value.replace('"', '""') + '"' if strict else value
    return to_text(value)


@function("VALUETOTEXT", kind="check")
def valuetotext(value: Scalar, form: Scalar = 0) -> str:
    mode = to_int(form)
    if mode not in (0, 1):
        raise FormulaError(VALUE)
    return _shown(value, strict=mode == 1)


@function("ARRAYTOTEXT", array=(0,))
def arraytotext(array: Value, form: Scalar = 0) -> str:
    mode = to_int(scalar(form))
    if mode not in (0, 1):
        raise FormulaError(VALUE)
    grid = as_grid(array)
    strict = mode == 1
    joiner = "," if strict else ", "
    lines = [joiner.join(_shown(item, strict) for item in row) for row in grid.rows]  # pyright: ignore[reportArgumentType]
    return ("{" + ";".join(lines) + "}") if strict else ", ".join(lines)
