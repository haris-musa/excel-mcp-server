"""Re-pointing references when a sheet is copied, as Excel does."""

from dataclasses import dataclass

from openpyxl.utils import quote_sheetname

from excel_mcp.formulas import split_top_level, unquote
from excel_mcp.operations.tables import is_valid_name
from excel_mcp.rewrite import ReferenceRewriter


@dataclass
class SheetCopyRefs(ReferenceRewriter):
    """References to the source sheet or its tables become references to the copy's."""

    source: str
    target: str
    tables: dict[str, str]
    """Lower-case source table name to the copy's table name."""

    def reference(self, text: str, host: str) -> str:
        parts = split_top_level(text, "!")
        if len(parts) > 1 and unquote(parts[0]).casefold() == self.source.casefold():
            return "!".join([quote_sheetname(self.target), *parts[1:]])
        table, bracket, rest = text.partition("[")
        if bracket and table.casefold() in self.tables:
            return self.tables[table.casefold()] + bracket + rest
        return text


@dataclass
class SheetRenameRefs(ReferenceRewriter):
    """References to a renamed sheet name the new one, quoted as it needs."""

    old: str
    new: str

    def reference(self, text: str, host: str) -> str:
        pieces = split_top_level(text, "!")
        if len(pieces) != 2:
            return text
        names = unquote(pieces[0]).split(":")
        if self.old.casefold() not in (name.casefold() for name in names):
            return text
        renamed = [self.new if name.casefold() == self.old.casefold() else name for name in names]
        return f"{_qualifier(':'.join(renamed))}!{pieces[1]}"


def _qualifier(names: str) -> str:
    """The sheet part of a reference: quoted only when Excel would quote it."""
    if is_valid_name(names.replace(":", "_")):
        return names
    return "'" + names.replace("'", "''") + "'"
