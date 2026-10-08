"""Moving references when rows or columns are inserted or deleted, as Excel does."""

import re
from dataclasses import dataclass, replace

from openpyxl.utils.cell import column_index_from_string, get_column_letter

from excel_mcp.formulas import split_top_level, unquote
from excel_mcp.package.lines import LineEdit
from excel_mcp.refs import MAX_COLUMN, MAX_ROW
from excel_mcp.rewrite import ReferenceRewriter

REF_ERROR = "#REF!"

_CELL = re.compile(r"(\$?)([A-Za-z]{1,3})(\$?)([0-9]+)")
_COLUMN = re.compile(r"(\$?)([A-Za-z]{1,3})")
_ROW = re.compile(r"(\$?)([0-9]+)")
_COLUMN_NAME = re.compile(r"\[([^\[\]#@][^\[\]]*)\]")


@dataclass(frozen=True)
class _Part:
    """One end of a reference: a cell, a whole column or a whole row, with its ``$`` markers."""

    column: int | None
    column_fixed: bool
    row: int | None
    row_fixed: bool

    def moved(self, rows: int, columns: int) -> "_Part | None":
        """The part after copying its formula by some lines, or None if that leaves the sheet."""
        column = self.column if self.column is None or self.column_fixed else self.column + columns
        row = self.row if self.row is None or self.row_fixed else self.row + rows
        if (column is not None and not 1 <= column <= MAX_COLUMN) or (
            row is not None and not 1 <= row <= MAX_ROW
        ):
            return None
        return replace(self, column=column, row=row)

    def text(self) -> str:
        column = ""
        if self.column is not None:
            column = ("$" if self.column_fixed else "") + get_column_letter(self.column)
        row = "" if self.row is None else ("$" if self.row_fixed else "") + str(self.row)
        return column + row


class Shifter(ReferenceRewriter):
    """Rewrites references for one edit. Like Excel, it leaves 3D references alone.

    ``dead_tables`` maps a lower-case table name to the lower-case names of its deleted
    columns, or to None when the whole table is deleted: references to them become #REF!.
    """

    def __init__(self, edit: LineEdit, dead_tables: dict[str, frozenset[str] | None]) -> None:
        self.edit = edit
        self.dead_tables = dead_tables

    def reference(self, text: str, host: str) -> str:
        table, bracket, _ = text.partition("[")
        if bracket and table.casefold() in self.dead_tables:
            dead = self.dead_tables[table.casefold()]
            columns = {name.casefold() for name in _COLUMN_NAME.findall(text)}
            return REF_ERROR if dead is None or columns & dead else text
        target = self.target(text, host)
        if target is None:
            return text
        prefix, parts = target
        moved = self._move(parts[0], parts[-1])
        if moved is None:
            return prefix + REF_ERROR
        return prefix + ":".join(part.text() for part in moved[: len(parts)])

    def target(self, text: str, host: str) -> tuple[str, list[_Part]] | None:
        """The ``Sheet!`` prefix and the cells of a reference to the edited sheet."""
        split = _split_reference(text)
        if split is None:
            return None
        qualifier, cells = split
        if (
            host if qualifier is None else unquote(qualifier)
        ).casefold() != self.edit.sheet.casefold():
            return None
        return ("" if qualifier is None else f"{qualifier}!"), cells

    def position(self, part: _Part) -> tuple[int | None, bool]:
        """The row or column an end of a reference sits in, and whether it is fixed with $."""
        if self.edit.axis == "rows":
            return part.row, part.row_fixed
        return part.column, part.column_fixed

    def _move(self, first: _Part, last: _Part) -> tuple[_Part, _Part] | None:
        rows = self.edit.axis == "rows"
        a, b = (first.row, last.row) if rows else (first.column, last.column)
        if a is None or b is None:
            return first, last
        if a > b:
            first, last, a, b = last, first, b, a
        moved = self.edit.span(a, b)
        if moved is None:
            return None
        if rows:
            return replace(first, row=moved[0]), replace(last, row=moved[1])
        return replace(first, column=moved[0]), replace(last, column=moved[1])


class Mover(ReferenceRewriter):
    """Rewrites a formula as Excel does when it is copied: only relative references move."""

    def __init__(self, rows: int, columns: int) -> None:
        self.rows = rows
        self.columns = columns

    def reference(self, text: str, host: str) -> str:
        split = _split_reference(text)
        if split is None:
            return text
        qualifier, cells = split
        moved = [cell.moved(self.rows, self.columns) for cell in cells]
        prefix = "" if qualifier is None else f"{qualifier}!"
        if None in moved:
            return prefix + REF_ERROR
        return prefix + ":".join(cell.text() for cell in moved if cell)


def _split_reference(text: str) -> tuple[str | None, list[_Part]] | None:
    """The sheet qualifier and the cells of a plain reference; None for names and the like."""
    pieces = split_top_level(text, "!")
    if len(pieces) > 2:
        return None
    parts = [_parse_part(part) for part in split_top_level(pieces[-1], ":")]
    cells = [part for part in parts if part is not None]
    if len(cells) != len(parts) or len(cells) > 2:
        return None
    if len(cells) == 1 and cells[0].row is None:
        return None  # A bare "Tax" is a name, not column Tax.
    return (pieces[0] if len(pieces) == 2 else None), cells


def _parse_part(text: str) -> _Part | None:
    if match := _CELL.fullmatch(text):
        column = column_index_from_string(match[2].upper())
        row = int(match[4])
        if column <= MAX_COLUMN and 1 <= row <= MAX_ROW:
            return _Part(column, match[1] == "$", row, match[3] == "$")
    elif match := _COLUMN.fullmatch(text):
        column = column_index_from_string(match[2].upper())
        if column <= MAX_COLUMN:
            return _Part(column, match[1] == "$", None, False)
    elif (match := _ROW.fullmatch(text)) and 1 <= int(match[2]) <= MAX_ROW:
        return _Part(None, False, int(match[2]), match[1] == "$")
    return None
