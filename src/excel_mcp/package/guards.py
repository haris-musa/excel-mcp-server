"""Edits that would leave preserved content in a state this server cannot store."""

import re
from typing import cast

from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from excel_mcp.errors import InvalidArgumentError
from excel_mcp.package.consistency import PIVOT_ENTRY, slicer_caches, used_caches
from excel_mcp.package.model import Part, state_of
from excel_mcp.package.scan import attributes

_TABLE_ENTRY = re.compile(r"<(?:[\w.-]+:)?tableSlicerCache\b[^>]*/>")


def check_pivot_removal(sheet: Worksheet, name: str) -> None:
    """Refuse deleting a PivotTable that slicers or timelines are connected to."""
    _refuse(_users(sheet, None, name))


def check_sheet_removal(sheet: Worksheet) -> None:
    """Refuse deleting a sheet whose PivotTables or tables have slicers on other sheets."""
    _refuse(_users(sheet, sheet, None))


def _users(sheet: Worksheet, leaving: Worksheet | None, pivot: str | None) -> list[str]:
    """Slicers and timelines that would be left connected to something that is deleted."""
    workbook = cast(Workbook, sheet.parent)
    state = state_of(workbook)
    sheets = {identifier: s for s, identifier in state.sheet_ids.items()}
    tables = {identifier: table for table, identifier in state.table_ids.items()}
    staying = used_caches(workbook, state, leaving)
    users = []
    for cache, link in slicer_caches(state).items():
        if cache not in staying or not isinstance(link.target, Part):
            continue
        text = link.target.data.decode("utf-8")
        for found in PIVOT_ENTRY.findall(text):
            values = attributes(found)
            on_sheet = sheets.get(int(values["tabId"])) is sheet
            if on_sheet and (pivot is None or values["name"].casefold() == pivot.casefold()):
                users.append(cache)
        if leaving is not None:
            for found in _TABLE_ENTRY.findall(text):
                table = tables.get(int(attributes(found)["tableId"]))
                if table is not None and table in sheet.tables.values():
                    users.append(cache)
    return users


def _refuse(users: list[str]) -> None:
    if users:
        raise InvalidArgumentError(
            f"Slicers or timelines ({', '.join(sorted(set(users)))}) are connected to it. Excel "
            "would keep them disconnected, which this server cannot store. Delete those "
            "slicers in Excel first."
        )
