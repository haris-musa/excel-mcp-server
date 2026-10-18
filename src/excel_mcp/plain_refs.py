"""Formulas made of plain cell references, handled without tokenizing them.

A large sheet holds hundreds of thousands of formulas such as ``=B2*C2`` or
``=IF(D2>50,"big","small")``. Tokenizing each one to move its references or to check it costs
more than everything else an edit does. `sub` finds the references of such a formula with one
regular expression, and gives up (returns None) on anything it cannot be sure of: sheet
names, names with brackets, whole rows and columns, array constants, spill references and
the like, which the caller then handles with the tokenizer.

A reference is plain when it stands alone between operators, parentheses, commas or spaces:
``A1``, ``$A$1``, ``A1:B5``. Text in quotes is skipped.
"""

import re
from collections.abc import Callable
from functools import lru_cache

from openpyxl.utils import column_index_from_string, get_column_letter

from excel_mcp.refs import MAX_COLUMN, MAX_ROW

_CELL = r"(\$?)([A-Za-z]{1,3})(\$?)([0-9]+)"
_BEFORE = r"(?<![^=+\-*/^&<>,(%\s])"
_AFTER = r"(?![^=+\-*/^&<>,)%\s])"
_REFERENCE = rf"{_BEFORE}{_CELL}(?::{_CELL})?{_AFTER}"
_OTHER_CHARACTERS = r"[!\'\[\]{}#@\\|:?]"
_REFERENCES = re.compile(_REFERENCE)
_OTHER = re.compile(_OTHER_CHARACTERS)
_TOKENS = re.compile(rf'"(?:[^"]|"")*"|{_REFERENCE}|(?P<other>{_OTHER_CHARACTERS})')

Match = re.Match[str]


class Cell:
    """One end of a plain reference, with the ``$`` markers kept."""

    __slots__ = ("column", "column_fixed", "row", "row_fixed")

    def __init__(self, column: int, column_fixed: bool, row: int, row_fixed: bool) -> None:
        self.column, self.column_fixed = column, column_fixed
        self.row, self.row_fixed = row, row_fixed

    def text(self) -> str:
        return (
            f"{'$' if self.column_fixed else ''}{column_letters(self.column)}"
            f"{'$' if self.row_fixed else ''}{self.row}"
        )


@lru_cache(maxsize=2048)
def column_letters(column: int) -> str:
    return get_column_letter(column)


@lru_cache(maxsize=2048)
def _column_number(letters: str) -> int:
    return column_index_from_string(letters.upper())


def _cell(match: Match, offset: int) -> Cell | None:
    column = _column_number(match[offset + 2])
    row = int(match[offset + 4])
    if column > MAX_COLUMN or not 1 <= row <= MAX_ROW:
        return None
    return Cell(column, bool(match[offset + 1]), row, bool(match[offset + 3]))


def references(match: Match) -> list[Cell] | None:
    """The ends of the reference a `sub` callback is given; None if it is not a reference
    (a name such as ``ZZZ1``), which stays as it is."""
    first = _cell(match, 0)
    if first is None:
        return None
    if match[6] is None:
        return [first]
    last = _cell(match, 4)
    return None if last is None else [first, last]


def sub(formula: str, replace: Callable[[Match], str]) -> str | None:
    """``formula`` with each plain reference replaced by ``replace(match)``, or None if it holds
    anything else that could be a reference."""
    if '"' not in formula:
        # A colon that is left once the references are gone belongs to a whole row or column.
        if _OTHER.search(_REFERENCES.sub("", formula)):
            return None
        return _REFERENCES.sub(replace, formula)
    unsure = False

    def visit(match: Match) -> str:
        nonlocal unsure
        if match["other"] is not None:
            unsure = True
            return ""
        if match[2] is None:  # text in quotes
            return match[0]
        return replace(match)

    done = _TOKENS.sub(visit, formula)
    return None if unsure else done


def shape(formula: str) -> str | None:
    """The formula with every reference replaced by ``A1``: formulas that differ only in
    which cells they use have the same shape. None for a formula that is not plain."""
    return sub(formula, lambda _: "A1")
