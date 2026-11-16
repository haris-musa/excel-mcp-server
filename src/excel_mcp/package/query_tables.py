"""Keeping a query table part in step with the columns of the table that shows its data.

A table loaded from Power Query or another connection has a query table part with one field
per table column, each naming its column by id, and a hidden name for the range the data
fills. Excel cannot open a file in which these disagree with the table, so after columns are
added or removed they are rewritten as Excel would: a removed column's field is listed as
deleted, a new column gets a field bound to no data, and the name covers the table without
the unbound columns at its edges.
"""

import re
from dataclasses import dataclass, replace
from xml.sax.saxutils import quoteattr

from openpyxl.utils.cell import absolute_coordinate

from excel_mcp.package.model import Part
from excel_mcp.package.patch import Edit, splice
from excel_mcp.package.scan import attributes
from excel_mcp.refs import CellRange, parse_range

QUERY_TABLE = "application/vnd.openxmlformats-officedocument.spreadsheetml.queryTable+xml"


def _element(name: str) -> re.Pattern[bytes]:
    """An element ``name`` with its content, written with or without a namespace prefix."""
    tag = rb"<(?:[\w.-]+:)?" + name.encode()
    return re.compile(tag + rb"\b[^>]*?(?:/>|>.*?</(?:[\w.-]+:)?" + name.encode() + rb">)", re.S)


_TABLE = re.compile(rb"<(?:[\w.-]+:)?table\b[^>]*>")
_COLUMN = re.compile(rb"<(?:[\w.-]+:)?tableColumn\b[^>]*?/?>")
_QUERY_TABLE = re.compile(rb"<(?:[\w.-]+:)?queryTable\b[^>]*>")
_REFRESH = re.compile(rb"<(?:[\w.-]+:)?queryTableRefresh\b[^>]*>")
_FIELD = _element("queryTableField")
_FIELDS = _element("queryTableFields")
_DELETED = _element("queryTableDeletedFields")
_DELETED_FIELD = re.compile(rb"<(?:[\w.-]+:)?deletedField\b[^>]*>")
_FIELD_ID = re.compile(rb'\squeryTableFieldId="[^"]*"')
_TABLE_COLUMN_ID = re.compile(rb'\btableColumnId="[^"]*"')
_REFRESH_COUNTS = re.compile(rb'\s(?:nextId|unboundColumnsLeft|unboundColumnsRight)="[^"]*"')
_DEFINED_NAME = re.compile(r"(<definedName\b[^>]*>)([^<]*)(</definedName>)")


@dataclass
class QueryTable:
    """A query table part and its table as they are saved, and the range of its data."""

    part: Part
    table: bytes
    name: str
    area: CellRange


def reconcile(query_table: Part, table: bytes) -> QueryTable | None:
    """The query table with one field for each column of the table; None if it has no fields."""
    fields = {_attributes(m[0])["id"]: m[0] for m in _FIELD.finditer(query_table.data)}
    refresh = _REFRESH.search(query_table.data)
    root = _QUERY_TABLE.search(query_table.data)
    table_root = _TABLE.search(table)
    if refresh is None or root is None or table_root is None or not fields:
        return None
    next_id = int(_attributes(refresh[0]).get("nextId", 0)) or max(map(int, fields)) + 1
    kept: list[bytes] = []
    bound: list[bool] = []
    edits: list[Edit] = []
    for match in _COLUMN.finditer(table):
        column = _attributes(match[0])["id"]
        field = _attributes(match[0]).get("queryTableFieldId")
        if field in fields:
            text = fields.pop(field)
            kept.append(_TABLE_COLUMN_ID.sub(f'tableColumnId="{column}"'.encode(), text, count=1))
            bound.append(_attributes(text).get("dataBound") != "0")
            continue
        field = str(next_id)
        next_id += 1
        kept.append(
            f'<queryTableField id="{field}" dataBound="0" tableColumnId="{column}"/>'.encode()
        )
        bound.append(False)
        edits.append((match.start(), match.end(), _with_field(match[0], column, field)))
    left = bound.index(True) if True in bound else 0
    right = bound[::-1].index(True) if True in bound else 0
    data = _fields(query_table.data, kept, list(fields.values()))
    counts = {"nextId": next_id, "unboundColumnsLeft": left, "unboundColumnsRight": right}
    data = _REFRESH.sub(lambda m: _with_counts(m[0], counts), data, count=1)
    area = parse_range(_attributes(table_root[0])["ref"])
    return QueryTable(
        replace(query_table, data=data),
        splice(table, edits),
        _attributes(root[0])["name"],
        replace(area, min_col=area.min_col + left, max_col=area.max_col - right),
    )


def with_areas(workbook: str, areas: dict[tuple[int, str], CellRange]) -> str:
    """The workbook XML with the hidden name of each query table covering its data.

    ``areas`` maps (sheet index, name) to the range, as the sheet-level names are keyed.
    """

    def area(match: re.Match[str]) -> str:
        values = attributes(match[1])
        key = (int(values.get("localSheetId", -1)), values["name"])
        if key not in areas:
            return match[0]
        sheet = match[2].rpartition("!")[0]
        return f"{match[1]}{sheet}!{absolute_coordinate(str(areas[key]))}{match[3]}"

    return _DEFINED_NAME.sub(area, workbook)


def _with_field(tag: bytes, column: str, field: str) -> bytes:
    """A ``<tableColumn>`` start tag bound to the query table field ``field``.

    Excel also needs a ``uniqueName`` on every column of a query table; it uses the id.
    """
    tag = _FIELD_ID.sub(b"", tag)
    end = len(tag) - (2 if tag.endswith(b"/>") else 1)
    extra = "" if b" uniqueName=" in tag else f' uniqueName="{column}"'
    return tag[:end] + f'{extra} queryTableFieldId="{field}"'.encode() + tag[end:]


def _with_counts(tag: bytes, counts: dict[str, int]) -> bytes:
    tag = _REFRESH_COUNTS.sub(b"", tag)
    end = len(tag) - (2 if tag.endswith(b"/>") else 1)
    added = "".join(f' {key}="{value}"' for key, value in counts.items() if value)
    return tag[:end] + added.encode() + tag[end:]


def _fields(data: bytes, kept: list[bytes], removed: list[bytes]) -> bytes:
    found = _DELETED.search(data)
    deleted = [m[0] for m in _DELETED_FIELD.finditer(found[0])] if found else []
    for field in removed:
        values = _attributes(field)
        if values.get("dataBound") != "0" and "name" in values:
            deleted.append(f"<deletedField name={quoteattr(values['name'])}/>".encode())
    fields = b'<queryTableFields count="%d">%s</queryTableFields>' % (len(kept), b"".join(kept))
    if deleted:
        fields += b'<queryTableDeletedFields count="%d">%s</queryTableDeletedFields>' % (
            len(deleted),
            b"".join(deleted),
        )
    return _FIELDS.sub(lambda _: fields, _DELETED.sub(b"", data), count=1)


def _attributes(element: bytes) -> dict[str, str]:
    """The attributes of the start tag of ``element``."""
    return attributes(element[: element.index(b">") + 1].decode("utf-8"))
