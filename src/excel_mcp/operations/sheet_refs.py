"""Re-pointing references when a sheet is copied, as Excel does."""

from dataclasses import dataclass

from openpyxl.formula import Tokenizer
from openpyxl.formula.tokenizer import Token
from openpyxl.utils import quote_sheetname

from excel_mcp.formulas import split_top_level, unquote


@dataclass
class SheetCopyRefs:
    """References to the source sheet or its tables become references to the copy's."""

    source: str
    target: str
    tables: dict[str, str]
    """Lower-case source table name to the copy's table name."""

    def formula(self, formula: str) -> str:
        tokenizer = Tokenizer(formula)
        for token in tokenizer.items:
            if token.subtype == Token.RANGE:
                token.value = self._reference(token.value)
        return tokenizer.render()

    def operand(self, operand: str) -> str:
        """For rule operands and names, which are stored without the "="."""
        return self.formula(f"={operand}").removeprefix("=")

    def _reference(self, reference: str) -> str:
        parts = split_top_level(reference, "!")
        if len(parts) > 1 and unquote(parts[0]).casefold() == self.source.casefold():
            return "!".join([quote_sheetname(self.target), *parts[1:]])
        table, bracket, rest = reference.partition("[")
        if bracket and table.casefold() in self.tables:
            return self.tables[table.casefold()] + bracket + rest
        return reference
