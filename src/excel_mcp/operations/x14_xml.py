"""Markup shared by the Excel 2010 extensions this server writes: sparklines and the
conditional formats that only exist there."""

import re
from xml.sax.saxutils import quoteattr

from excel_mcp.operations.formatting import parse_color

X14 = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
XM = "http://schemas.microsoft.com/office/excel/2006/main"


def attributes(pairs: list[tuple[str, str | None]]) -> str:
    """`` name="value"`` for each pair, in order, leaving out those without a value."""
    return "".join(f" {name}={quoteattr(value)}" for name, value in pairs if value is not None)


def color(tag: str, value: str) -> str:
    return f'<x14:{tag} rgb="{parse_color(value)}"/>'


def number_text(value: float) -> str:
    return str(int(value)) if value == int(value) else repr(value)


_PLAIN_NAME = re.compile(r"[A-Za-z_][A-Za-z0-9_.]*")
_CELL_LIKE = re.compile(r"[A-Za-z]{1,3}\d+|R\d*C\d*", re.IGNORECASE)


def sheet_prefix(title: str) -> str:
    """The sheet name as Excel writes it in a reference: quoted only when it must be."""
    if _PLAIN_NAME.fullmatch(title) and not _CELL_LIKE.fullmatch(title):
        return title
    return "'" + title.replace("'", "''") + "'"
