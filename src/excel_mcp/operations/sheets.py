"""Worksheet management and structure: create, rename, copy, delete, insert and delete lines."""

from typing import Literal

from openpyxl.chartsheet import Chartsheet
from openpyxl.utils.cell import get_column_letter
from openpyxl.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from excel_mcp.errors import InvalidArgumentError
from excel_mcp.operations.cells import used_range
from excel_mcp.package.guards import check_sheet_removal
from excel_mcp.refs import MAX_COLUMN, MAX_ROW
from excel_mcp.workspace import get_sheet

Axis = Literal["rows", "columns"]

_INVALID_NAME_CHARACTERS = set("[]:*?/\\")


def validate_sheet_name(name: str, existing: list[str]) -> None:
    if not name.strip() or len(name) > 31:
        raise InvalidArgumentError("Sheet names must be 1 to 31 characters long.")
    if _INVALID_NAME_CHARACTERS & set(name):
        raise InvalidArgumentError("Sheet names cannot contain any of [ ] : * ? / \\.")
    if name.startswith("'") or name.endswith("'"):
        raise InvalidArgumentError("Sheet names cannot start or end with an apostrophe.")
    if name.casefold() in (other.casefold() for other in existing):
        raise InvalidArgumentError(f"A sheet named {name!r} already exists.")


def create_sheet(workbook: Workbook, name: str, position: int | None) -> Worksheet:
    validate_sheet_name(name, workbook.sheetnames)
    index = None if position is None else position - 1
    return workbook.create_sheet(name, index)


def rename_sheet(workbook: Workbook, name: str, new_name: str) -> None:
    sheet = get_sheet(workbook, name)
    validate_sheet_name(new_name, workbook.sheetnames)
    sheet.title = new_name


def delete_sheet(workbook: Workbook, name: str) -> None:
    if name in workbook.sheetnames and isinstance(workbook[name], Chartsheet):
        workbook.remove(workbook[name])
        return
    sheet = get_sheet(workbook, name)
    visible = [other for other in workbook.worksheets if other.sheet_state == "visible"]
    if visible == [sheet]:
        raise InvalidArgumentError("A workbook must keep at least one visible worksheet.")
    check_sheet_removal(sheet)
    workbook.remove(sheet)


def insert_lines(sheet: Worksheet, axis: Axis, at: int, count: int) -> None:
    area = used_range(sheet)
    if axis == "rows":
        _check_room(area.max_row + count, MAX_ROW, "rows")
        sheet.insert_rows(at, count)
    else:
        _check_room(area.max_col + count, MAX_COLUMN, "columns")
        sheet.insert_cols(at, count)


def delete_lines(sheet: Worksheet, axis: Axis, at: int, count: int) -> None:
    if axis == "rows":
        sheet.delete_rows(at, count)
    else:
        sheet.delete_cols(at, count)


def describe_lines(axis: Axis, at: int, count: int) -> str:
    """``1 row at row 2`` or ``3 columns at column C``."""
    unit = axis.removesuffix("s")
    position = at if axis == "rows" else get_column_letter(at)
    return f"{count} {axis if count != 1 else unit} at {unit} {position}"


def _check_room(needed: int, limit: int, axis: Axis) -> None:
    if needed > limit:
        raise InvalidArgumentError(
            f"Inserting would push data past the last of Excel's {limit:,} {axis}."
        )
