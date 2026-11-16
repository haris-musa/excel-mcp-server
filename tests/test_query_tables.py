"""A table loaded from Power Query keeps a query table part that matches its columns.

The fixture was saved by Microsoft Excel; see tests/fixtures/README.md. Each expected part
is what Excel itself writes for the same edit.
"""

import re
from pathlib import Path

import pytest
from openpyxl import load_workbook

from tests.conftest import ToolCall
from tests.package_support import (
    assert_package_is_consistent,
    copy_fixture,
    read_parts,
    relationships,
    sheet_part,
    text,
)

pytestmark = pytest.mark.anyio

BOOK = {"path": "book.xlsx", "sheet": "Stock"}
QUERY_TABLE = "xl/queryTables/queryTable1.xml"
TABLE = "xl/tables/table1.xml"


async def _edited(call: ToolCall, files: Path, tool: str, **arguments: object) -> dict[str, str]:
    copy_fixture(files, "excel_query_table.xlsx")
    await call(tool, **BOOK, **arguments)
    parts = read_parts(files / "book.xlsx")
    assert_package_is_consistent(parts)
    load_workbook(files / "book.xlsx").close()
    names = re.findall(r"<definedName\b[^>]*>([^<]*)<", text(parts, "xl/workbook.xml"))
    return {
        "fields": _fields(text(parts, QUERY_TABLE)),
        "refresh": re.findall(r"<queryTableRefresh\b[^>]*>", text(parts, QUERY_TABLE))[0],
        "deleted": " ".join(re.findall(r'<deletedField name="([^"]*)"', text(parts, QUERY_TABLE))),
        "columns": _columns(text(parts, TABLE)),
        "name": " ".join(names),
    }


def _fields(xml: str) -> str:
    fields = re.findall(r"<queryTableField\b([^>]*?)/?>", xml)
    return " | ".join(" ".join(f.split()) for f in fields)


def _columns(xml: str) -> str:
    columns = re.findall(r"<tableColumn\b([^>]*?)/?>", xml)
    return " | ".join(
        " ".join(sorted(re.findall(r'(?:id|uniqueName|queryTableFieldId)="[^"]*"', c)))
        for c in columns
    )


async def test_the_query_table_part_survives_an_edit(call: ToolCall, files: Path) -> None:
    copy_fixture(files, "excel_query_table.xlsx")
    await call("write_range", **BOOK, at="D1", rows=[["note"]])
    parts = read_parts(files / "book.xlsx")

    links = relationships(parts, TABLE)
    assert ("queryTable", QUERY_TABLE) in links.values()
    assert TABLE in [t for _, t in relationships(parts, sheet_part(parts, "Stock")).values()]
    assert 'tableType="queryTable"' in text(parts, TABLE)
    assert_package_is_consistent(parts)


async def test_an_inserted_column_gets_an_unbound_field(call: ToolCall, files: Path) -> None:
    found = await _edited(call, files, "insert_rows_or_columns", axis="columns", start=2, count=1)

    assert found["fields"] == (
        'id="1" name="Name" tableColumnId="3" | id="3" dataBound="0" tableColumnId="5"'
        ' | id="2" name="Qty" tableColumnId="4"'
    )
    assert found["columns"] == (
        'id="3" queryTableFieldId="1" uniqueName="3"'
        ' | id="5" queryTableFieldId="3" uniqueName="5"'
        ' | id="4" queryTableFieldId="2" uniqueName="4"'
    )
    assert 'nextId="4"' in found["refresh"]
    assert "unboundColumns" not in found["refresh"]
    assert found["name"] == "Stock!$A$1:$C$4"


async def test_a_deleted_column_is_listed_as_a_deleted_field(call: ToolCall, files: Path) -> None:
    found = await _edited(call, files, "delete_rows_or_columns", axis="columns", start=2, count=1)

    assert found["fields"] == 'id="1" name="Name" tableColumnId="3"'
    assert found["deleted"] == "Qty"
    assert found["name"] == "Stock!$A$1:$A$4"


async def test_columns_added_by_resizing_are_unbound_at_the_right(
    call: ToolCall, files: Path
) -> None:
    found = await _edited(call, files, "set_table", name="Stock", range="A1:D4")

    assert found["fields"].endswith(
        'id="3" dataBound="0" tableColumnId="5" | id="4" dataBound="0" tableColumnId="6"'
    )
    assert 'unboundColumnsRight="2"' in found["refresh"]
    assert found["name"] == "Stock!$A$1:$B$4"


async def test_the_name_shrinks_with_the_table(call: ToolCall, files: Path) -> None:
    found = await _edited(call, files, "set_table", name="Stock", range="A1:A4")

    assert found["deleted"] == "Qty"
    assert found["name"] == "Stock!$A$1:$A$4"


async def test_a_copied_query_table_is_a_plain_table(call: ToolCall, files: Path) -> None:
    copy_fixture(files, "excel_query_table.xlsx")
    await call("copy_sheet", path="book.xlsx", sheet="Stock", new_name="Copy")
    parts = read_parts(files / "book.xlsx")
    tables = {
        t
        for _, t in relationships(parts, sheet_part(parts, "Copy")).values()
        if t.startswith("xl/tables/")
    }

    assert len(tables) == 1
    copy = text(parts, tables.pop())
    assert "queryTable" not in copy
    assert ("queryTable", QUERY_TABLE) in relationships(parts, TABLE).values()
    assert_package_is_consistent(parts)
