"""Rows of a large sheet as cells and back, without an XML library.

openpyxl reads and writes a cell by building and parsing XML elements, which costs more than
anything else an edit of a large range does. The cells of the forms Excel and openpyxl write are
read and written here as text, with the same results. A row of any other form (rich text, array
formulas, attributes that are not usual) is left to openpyxl: `read` gives None for it, and
`render` leaves it out.
"""

# pyright: reportAttributeAccessIssue=false, reportPrivateUsage=false
# openpyxl's stubs lack the members of cells that it writes from.

import re
import warnings
from collections.abc import Mapping
from dataclasses import dataclass
from datetime import datetime
from html import unescape
from typing import Any
from xml.sax.saxutils import escape

from openpyxl.cell._writer import _set_attributes
from openpyxl.cell.rich_text import CellRichText
from openpyxl.compat import safe_string
from openpyxl.formula.translate import Translator
from openpyxl.utils import column_index_from_string
from openpyxl.utils.datetime import from_excel, from_ISO8601
from openpyxl.worksheet.formula import ArrayFormula, DataTableFormula

_ROW = re.compile(rb'<row\b[^>]*?\br="(\d+)"[^>]*?(?:/>|>(.*)</row>)', re.S)
_CELL = re.compile(
    rb'<c r="(?P<column>[A-Z]+)(?P<row>\d+)"(?: s="(?P<style>\d+)")?(?: t="(?P<type>\w+)")?'
    rb"(?: ?/>|>"
    rb'(?P<formula><f(?: t="(?P<kind>\w+)")?(?: ref="[^"]*")?(?: si="(?P<group>\d+)")?'
    rb"(?: ?/>|>(?P<text>[^<]*)</f>))?"
    rb"(?:<v>(?P<value>[^<]*)</v>|<v ?/>)?"
    rb'(?:<is><t(?: xml:space="preserve")?>(?P<inline>[^<]*)</t></is>)?'
    rb"</c>)"
)

Item = tuple[int, int, Any, str]
"""A cell: its column, style, value and type."""


@dataclass
class Reading:
    """What reading cells needs of the workbook."""

    strings: list[str]
    shared: dict[str, Translator]
    date_formats: set[int]
    timedelta_formats: set[int]
    epoch: datetime
    data_only: bool


def read(row: bytes, reading: Reading) -> tuple[int, list[Item]] | None:
    """The row number and cells of a row of XML; None if it is not in a form read here."""
    found = _ROW.search(row)
    if found is None:
        return None
    number = int(found[1])
    cells = found[2] or b""
    items: list[Item] = []
    position = 0
    for cell in _CELL.finditer(cells):
        if cell.start() != position or int(cell["row"]) != number:
            return None
        position = cell.end()
        item = _item(cell, number, reading)
        if item is None:
            return None
        items.append(item)
    return (number, items) if position == len(cells) else None


def _item(cell: re.Match[bytes], row: int, reading: Reading) -> Item | None:
    kind = cell["type"]
    data_type = kind.decode() if kind else "n"
    style = int(cell["style"]) if cell["style"] else 0
    letters = cell["column"].decode()
    value = None if data_type == "inlineStr" else (_text(cell["value"]) or None)
    if cell["formula"] is not None and not reading.data_only:
        if cell["kind"] not in (None, b"normal", b"shared"):
            return None
        return (
            column_index_from_string(letters),
            style,
            _formula(cell, f"{letters}{row}", reading),
            "f",
        )
    if value is not None:
        if data_type == "n":
            value = float(value) if "." in value or "e" in value.lower() else int(value)
            if style in reading.date_formats:
                data_type = "d"
                try:
                    value = from_excel(
                        value, reading.epoch, timedelta=style in reading.timedelta_formats
                    )
                except (OverflowError, ValueError):
                    warnings.warn(
                        f"Cell {letters}{row} is marked as a date but the serial value {value} "
                        "is outside the limits for dates. The cell will be treated as an error.",
                        stacklevel=2,
                    )
                    data_type, value = "e", "#VALUE!"
        elif data_type == "s":
            value = reading.strings[int(value)]
        elif data_type == "b":
            value = bool(int(value))
        elif data_type == "str":
            data_type = "s"
        elif data_type == "d":
            value = from_ISO8601(value)
    elif data_type == "inlineStr" and cell["inline"] is not None:
        data_type, value = "s", _text(cell["inline"])
    return column_index_from_string(letters), style, value, data_type


def _formula(cell: re.Match[bytes], coordinate: str, reading: Reading) -> str:
    value = "=" + _text(cell["text"])
    if cell["kind"] == b"shared":
        group = cell["group"].decode()
        if group in reading.shared:
            return reading.shared[group].translate_formula(coordinate)
        if value != "=":
            reading.shared[group] = Translator(value, coordinate)
    return value


def _text(text: bytes | None) -> str:
    if text is None:
        return ""
    if b"\r" in text:  # An XML parser reads every kind of line break as "\n".
        text = text.replace(b"\r\n", b"\n").replace(b"\r", b"\n")
    return unescape(text.decode()) if b"&" in text else text.decode()


def render(
    cells: Mapping[tuple[int, int], Any], dimensions: Mapping[int, Any]
) -> tuple[dict[int, bytes], dict[tuple[int, int], Any]]:
    """The rows of ``cells`` as XML, and the cells that are left for openpyxl to write.

    A row is left to openpyxl if one of its cells has a link, a note or a value written in a
    form that is not handled here. ``dimensions`` give the attributes of rows (their height).
    """
    by_row: dict[int, list[Any]] = {}
    for (row, _), cell in sorted(cells.items()):
        by_row.setdefault(row, []).append(cell)
    rendered: dict[int, bytes] = {}
    rest: dict[tuple[int, int], Any] = {}
    for row, row_cells in by_row.items():
        xml = _row(row, row_cells, dimensions.get(row))
        if xml is None:
            rest.update({(row, cell.column): cell for cell in row_cells})
        else:
            rendered[row] = xml
    for row in dimensions.keys() - by_row.keys():
        rendered[row] = _row(row, [], dimensions[row]) or b""
    return rendered, rest


def _row(number: int, cells: list[Any], dimension: Any) -> bytes | None:
    attributes = {"r": str(number), **(dict(dimension) if dimension is not None else {})}
    opening = "".join(
        f' {name}="{escape(value, {chr(34): "&quot;"})}"' for name, value in attributes.items()
    )
    parts = [f"<row{opening}>"]
    for cell in cells:
        if cell._value is None and not cell.has_style and not cell._comment:
            continue
        xml = _cell(cell)
        if xml is None:
            return None
        parts.append(xml)
    parts.append("</row>")
    return "".join(parts).encode()


def _cell(cell: Any) -> str | None:
    value = cell._value
    if (
        cell._comment is not None
        or cell.hyperlink is not None
        or isinstance(value, ArrayFormula | DataTableFormula | CellRichText)
    ):
        return None
    value, attributes = _set_attributes(cell, cell.has_style)
    opening = "<c" + "".join(f' {name}="{text}"' for name, text in attributes.items())
    if value is None or value == "":
        return opening + "/>"
    if cell.data_type == "f":
        return f"{opening}><f>{escape(value[1:])}</f><v /></c>"
    if cell.data_type == "s":
        space = ' xml:space="preserve"' if value.strip() and value != value.strip() else ""
        return f"{opening}><is><t{space}>{escape(value)}</t></is></c>"
    return f"{opening}><v>{escape(safe_string(value))}</v></c>"
