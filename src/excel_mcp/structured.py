"""Structured table references: ``Sales[Price]``, ``[@Price]``, ``Sales[[#Totals],[Price]]``."""

import re
from dataclasses import dataclass

from openpyxl.formula import Tokenizer
from openpyxl.formula.tokenizer import Token

ITEMS = {
    "#all": "#All",
    "#data": "#Data",
    "#headers": "#Headers",
    "#totals": "#Totals",
    "#this row": "#This Row",
}
_ESCAPED = re.compile(r"'(.)")
_SPECIAL = re.compile(r"([\[\]#'])")
_REFERENCE = re.compile(r"([^\[\]]*)(\[.*\])", re.S)


@dataclass(frozen=True)
class StructuredRef:
    """``items`` are canonical ('#Data'); the columns run from ``first`` to ``last``."""

    table: str | None
    items: tuple[str, ...]
    first: str | None
    last: str | None
    shorthand: bool = False  # written with '@'

    def long_form(self, table: str) -> str:
        """The reference as Excel stores it in a file, for a table of this name."""
        parts = [f"[{item}]" for item in self.items]
        if self.first is not None:
            first = f"[{_escape(self.first)}]"
            parts.append(first if self.last is None else f"{first}:[{_escape(self.last)}]")
        if not parts:
            return f"{table}[#Data]"
        return table + (parts[0] if len(parts) == 1 else f"[{','.join(parts)}]")


def parse_structured(text: str) -> StructuredRef | None:
    """None when the text is not a structured reference."""
    match = _REFERENCE.fullmatch(text)
    if not match:
        return None
    inner = match[2][1:-1]
    shorthand = inner.startswith("@")
    items = ["#This Row"] if shorthand else []
    columns: list[str] = []
    if shorthand:
        inner = inner[1:]
    if inner.startswith("[") and inner.endswith("]"):
        if not all(_part(part.strip(), items, columns) for part in _split(inner, ",")):
            return None
    elif inner.casefold() in ITEMS:
        items.append(ITEMS[inner.casefold()])
    elif inner:
        columns.append(_unescape(inner))
    if len(columns) > 2:
        return None
    last = columns[1] if len(columns) == 2 else None
    return StructuredRef(
        match[1] or None, tuple(items), columns[0] if columns else None, last, shorthand
    )


def qualify_formula(formula: str, table: str) -> str:
    """The formula with the references to the table's own parts that are written without its
    name (``[Price]``, ``[@Price]``) or with ``@`` qualified as Excel stores them."""
    tokenizer = Tokenizer(formula)
    for token in tokenizer.items:
        if token.type == Token.OPERAND and token.subtype == Token.RANGE:
            found = parse_structured(token.value)
            if found and (found.table is None or found.shorthand):
                token.value = found.long_form(found.table or table)
    return tokenizer.render()


def _part(part: str, items: list[str], columns: list[str]) -> bool:
    pieces = _split(part, ":")
    if len(pieces) > 2 or not all(p.startswith("[") and p.endswith("]") for p in pieces):
        return False
    names = [p[1:-1] for p in pieces]
    if len(names) == 1 and names[0].casefold() in ITEMS:
        items.append(ITEMS[names[0].casefold()])
    else:
        columns.extend(_unescape(name) for name in names)
    return True


def _split(text: str, separator: str) -> list[str]:
    """Split at the separator outside brackets; ``'`` escapes the next character."""
    parts, depth, start, index = [], 0, 0, 0
    while index < len(text):
        character = text[index]
        if character == "'":
            index += 1
        elif character == "[":
            depth += 1
        elif character == "]":
            depth -= 1
        elif character == separator and depth == 0:
            parts.append(text[start:index])
            start = index + 1
        index += 1
    return [*parts, text[start:]]


def _unescape(name: str) -> str:
    return _ESCAPED.sub(r"\1", name)


def _escape(name: str) -> str:
    return _SPECIAL.sub(r"'\1", name)
