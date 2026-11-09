"""Translating a shared formula to the cells that use it, tokenizing it once.

openpyxl translates a shared formula for each cell by taking its master apart again and
moving the references, so a long formula filled over many cells costs its length for every
cell. `SharedTranslator` takes the master apart once, into the text between its references and
the references' relative rows and columns, so a cell costs only its references and a join.
The result is what openpyxl's `Translator` gives, for every formula and every cell.
"""

from functools import lru_cache

import openpyxl.worksheet._reader as sheet_reader
from openpyxl.formula.tokenizer import Token
from openpyxl.formula.translate import Translator, TranslatorError
from openpyxl.utils import column_index_from_string, coordinate_to_tuple, get_column_letter

Moving = tuple[bool, int]
"""A relative column (True) or row (False) of a reference, as its number."""


class SharedTranslator(Translator):
    row: int
    col: int

    def __init__(self, formula: str, origin: str) -> None:
        super().__init__(formula, origin)
        self._constant: str | None = None
        items: list[str | Moving] = []
        tokens = self.tokenizer.items
        if not tokens:
            self._constant = ""
        elif tokens[0].type == Token.LITERAL:
            self._constant = tokens[0].value
        else:
            items.append("=")
            for token in tokens:
                if token.type == Token.OPERAND and token.subtype == Token.RANGE:
                    self._split_range(token.value, items)
                else:
                    items.append(token.value)
        # The text before each moving part and after the last: one more than there are parts.
        self._texts: list[str] = []
        self._moving: list[Moving] = []
        run: list[str] = []
        for item in items:
            if isinstance(item, str):
                run.append(item)
            else:
                self._texts.append("".join(run))
                self._moving.append(item)
                run = []
        self._texts.append("".join(run))
        self._first_row = min((n for column, n in self._moving if not column), default=1)
        self._columns_at: int | None = None
        self._folded: tuple[list[str], list[int]] = ([], [])

    def _split_range(self, reference: str, items: list[str | Moving]) -> None:
        """Split a reference as `Translator.translate_range` reads it."""
        sheet, reference = self.strip_ws_name(reference)
        if not self.CELL_REF_RE.match(reference) and ":" not in reference:
            items.append(reference)  # a name: openpyxl drops its sheet, so the output must too
            return
        items.append(sheet)
        if match := self.ROW_RANGE_RE.match(reference):
            _part(match[1], False, items)
            items.append(":")
            _part(match[2], False, items)
        elif match := self.COL_RANGE_RE.match(reference):
            _part(match[1], True, items)
            items.append(":")
            _part(match[2], True, items)
        elif ":" in reference:
            for number, piece in enumerate(reference.split(":")):
                if number:
                    items.append(":")
                self._split_range(piece, items)
        elif match := self.CELL_REF_RE.match(reference):
            _part(match[1], True, items)
            _part(match[2], False, items)

    def translate_formula(
        self, dest: str | None = None, row_delta: int = 0, col_delta: int = 0
    ) -> str:
        if self._constant is not None:
            return self._constant
        if dest:
            row, column = coordinate_to_tuple(dest)
            row_delta, col_delta = row - self.row, column - self.col
        texts, rows = self._with_columns(col_delta)
        if rows and self._first_row + row_delta <= 0:
            raise TranslatorError("Formula out of range")
        out = [texts[0]]
        for number, text in zip(rows, texts[1:], strict=True):
            out.append(str(number + row_delta))
            out.append(text)
        return "".join(out)

    def _with_columns(self, delta: int) -> tuple[list[str], list[int]]:
        """The texts around the rows of the references, with the columns moved by ``delta``.
        A formula filled down moves them all by the same, so the last one is kept."""
        if self._columns_at != delta:
            texts: list[str] = []
            rows: list[int] = []
            run = [self._texts[0]]
            for (column, number), text in zip(self._moving, self._texts[1:], strict=True):
                if column:
                    run += (_letters(number + delta), text)
                else:
                    rows.append(number)
                    texts.append("".join(run))
                    run = [text]
            texts.append("".join(run))
            self._folded, self._columns_at = (texts, rows), delta
        return self._folded


def _part(text: str, column: bool, items: list[str | Moving]) -> None:
    if text.startswith("$"):
        items.append(text)
    else:
        items.append((column, column_index_from_string(text) if column else int(text)))


@lru_cache(maxsize=4096)
def _letters(number: int) -> str:
    try:
        return get_column_letter(number)
    except ValueError:
        raise TranslatorError("Formula out of range") from None


def install() -> None:
    """Make openpyxl translate the shared formulas of the sheets it reads with it. Safe to
    call again."""
    sheet_reader.Translator = SharedTranslator  # pyright: ignore[reportAttributeAccessIssue]
