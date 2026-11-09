"""Find and replace in cell contents, as Excel's Replace does."""

import re

from openpyxl.worksheet.worksheet import Worksheet
from pydantic import BaseModel

from excel_mcp import lazy_find
from excel_mcp.calc.values import MAX_TEXT
from excel_mcp.config import Limits
from excel_mcp.errors import LimitExceededError
from excel_mcp.lazy_workbook import LazyCells
from excel_mcp.operations.cells import parsed_cells, store_typed
from excel_mcp.text_budget import TextBudget
from excel_mcp.workspace import sheet_names
from excel_mcp.xlfn import display_formula


class ReplaceResult(BaseModel):
    replaced: dict[str, int]


def replace_cells(
    sheets: list[Worksheet],
    query: str,
    replacement: str,
    *,
    exact: bool,
    case_sensitive: bool,
    in_formulas: bool,
    limits: Limits,
) -> ReplaceResult:
    """Replace text in text, number and (optionally) formula cells; the result is retyped.

    Excel searches the text of a formula, not its result, so formulas are replaced in place.
    The result goes through the same rules as typing it in: it may become a number, and a
    formula must pass the formula check. Dates, booleans and array formulas are left alone.
    """
    flags = 0 if case_sensitive else re.IGNORECASE
    pattern = re.compile(re.escape(query), flags)
    found = pattern.fullmatch if exact else pattern.search

    def matches(text: str) -> bool:
        return found(text) is not None

    grown = TextBudget(
        limits.max_cells,
        "Replacing would add more than {limit} characters of text. Replace in one sheet at a "
        "time, or with shorter text.",
    )
    # A workbook can hold one long text and point millions of cells at it: each text is worked
    # out once, and its result is shared by every cell that holds it.
    results: dict[str, str | None] = {}
    replaced: dict[str, int] = {}
    for sheet in sheets:
        names = sheet_names(sheet)
        changed = 0
        if isinstance(sheet._cells, LazyCells):
            lazy_find.load_matches(sheet._cells, matches, query, in_formulas=in_formulas)
        for cell in list(parsed_cells(sheet)):
            text = _searchable_text(cell.value, cell.data_type, in_formulas)
            if text is None:
                continue
            if text not in results:
                new_text = _replaced(pattern, text, query, replacement, exact=exact)
                grown.spend(max(0, len(new_text) - len(text)))
                results[text] = None if new_text == text else new_text
            if (new_text := results[text]) is not None:
                store_typed(cell, new_text, names)
                changed += 1
        if changed:
            replaced[sheet.title] = changed
    return ReplaceResult(replaced=replaced)


def _replaced(
    pattern: re.Pattern[str], text: str, query: str, replacement: str, *, exact: bool
) -> str:
    """``text`` with the matches replaced; refused, before it is built, if too long for a cell."""
    if exact:
        if pattern.fullmatch(text) is None:
            return text
        _fit(len(replacement))
        return replacement
    if len(replacement) > len(query):
        found = list(pattern.finditer(text))
        _fit(len(text) + sum(len(replacement) - f.end() + f.start() for f in found))
    return pattern.sub(lambda _: replacement, text)


def _fit(size: int) -> None:
    if size > MAX_TEXT:
        raise LimitExceededError(
            f"Replacing would make a cell longer than {MAX_TEXT:,} characters, the limit of "
            "an Excel cell."
        )


def _searchable_text(value: object, data_type: str, in_formulas: bool) -> str | None:
    if data_type == "f":
        return display_formula(value) if in_formulas and isinstance(value, str) else None
    if data_type == "s":
        return str(value)
    if data_type == "n":
        return str(int(value)) if isinstance(value, float) and value == int(value) else str(value)
    return None
