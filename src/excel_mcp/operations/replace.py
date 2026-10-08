"""Find and replace in cell contents, as Excel's Replace does."""

import re

from openpyxl.worksheet.worksheet import Worksheet
from pydantic import BaseModel

from excel_mcp.operations.cells import store_typed, stored_cells
from excel_mcp.workspace import sheet_names
from excel_mcp.xlfn import remove_prefixes


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
) -> ReplaceResult:
    """Replace text in text, number and (optionally) formula cells; the result is retyped.

    Excel searches the text of a formula, not its result, so formulas are replaced in place.
    The result goes through the same rules as typing it in: it may become a number, and a
    formula must pass the formula check. Dates, booleans and array formulas are left alone.
    """
    flags = 0 if case_sensitive else re.IGNORECASE
    pattern = re.compile(re.escape(query), flags)
    replaced: dict[str, int] = {}
    for sheet in sheets:
        names = sheet_names(sheet)
        changed = 0
        for cell in list(stored_cells(sheet)):
            text = _searchable_text(cell.value, cell.data_type, in_formulas)
            if text is None:
                continue
            if exact:
                new_text = replacement if pattern.fullmatch(text) else text
            else:
                new_text = pattern.sub(lambda _: replacement, text)
            if new_text != text:
                store_typed(cell, new_text, names)
                changed += 1
        if changed:
            replaced[sheet.title] = changed
    return ReplaceResult(replaced=replaced)


def _searchable_text(value: object, data_type: str, in_formulas: bool) -> str | None:
    if data_type == "f":
        return remove_prefixes(value) if in_formulas and isinstance(value, str) else None
    if data_type == "s":
        return str(value)
    if data_type == "n":
        return str(int(value)) if isinstance(value, float) and value == int(value) else str(value)
    return None
