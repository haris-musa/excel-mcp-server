"""Worksheet management and structure: create, rename, copy, delete, insert and delete lines."""

from openpyxl.chartsheet import Chartsheet
from openpyxl.utils.cell import get_column_letter
from openpyxl.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from excel_mcp.errors import InvalidArgumentError
from excel_mcp.operations.pivot_index import workbook_pivots
from excel_mcp.operations.sheet_refs import SheetRenameRefs
from excel_mcp.operations.workbook_rewrite import (
    gated_operand,
    rewrite_charts,
    rewrite_formulas,
    rewrite_names,
    rewrite_rules,
)
from excel_mcp.package.guards import check_sheet_removal
from excel_mcp.package.lines import Axis
from excel_mcp.package.references import rewrite_formulas as rewrite_preserved
from excel_mcp.workspace import get_sheet

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
    old_name = sheet.title
    sheet.title = new_name
    refs = SheetRenameRefs(old_name, new_name)
    names = workbook.sheetnames
    rewrite_formulas(workbook, refs, names)
    rewrite_names(workbook, refs, names)
    rewrite_rules(workbook, refs, names)
    rewrite_charts(workbook, refs, names)
    rewrite_preserved(workbook, lambda text, host: gated_operand(text, host, refs, names))
    for pivot in workbook_pivots(workbook):
        source = pivot.cache.cacheSource.worksheetSource
        if source is not None and (source.sheet or "").casefold() == old_name.casefold():
            source.sheet = new_name


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


def describe_lines(axis: Axis, at: int, count: int) -> str:
    """``1 row at row 2`` or ``3 columns at column C``."""
    unit = axis.removesuffix("s")
    position = at if axis == "rows" else get_column_letter(at)
    return f"{count} {axis if count != 1 else unit} at {unit} {position}"
