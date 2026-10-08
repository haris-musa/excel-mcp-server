"""Defined names: workbook- or sheet-scoped labels for ranges and constants."""

from openpyxl.workbook import Workbook
from openpyxl.workbook.defined_name import DefinedName, DefinedNameDict
from openpyxl.worksheet.worksheet import Worksheet
from pydantic import BaseModel

from excel_mcp.errors import InvalidArgumentError
from excel_mcp.formulas import check_formula
from excel_mcp.operations.tables import is_valid_name, table_names


class DefinedNameInfo(BaseModel):
    name: str
    refers_to: str
    sheet: str | None = None


def list_defined_names(workbook: Workbook) -> list[DefinedNameInfo]:
    """Workbook-scoped names first, then each sheet's own names."""
    scopes: list[tuple[str | None, DefinedNameDict]] = [(None, workbook.defined_names)]
    scopes += [(sheet.title, sheet.defined_names) for sheet in workbook.worksheets]
    return [
        DefinedNameInfo(name=name, refers_to=defined.attr_text or "", sheet=scope)
        for scope, names in scopes
        for name, defined in sorted(names.items())
    ]


def set_defined_name(
    workbook: Workbook, name: str, refers_to: str, sheet: Worksheet | None
) -> bool:
    """Create or replace a name. Returns True when it replaced an existing one."""
    if not is_valid_name(name) or name.casefold().startswith("_xlnm"):
        raise InvalidArgumentError(
            f"Invalid name {name!r}. Start with a letter or underscore and use only "
            "letters, digits, underscores and periods; it cannot look like a cell "
            "reference such as 'A1'."
        )
    if name.casefold() in table_names(workbook):
        raise InvalidArgumentError(f"The name {name!r} is already used by a table.")
    reference = refers_to.removeprefix("=").strip()
    if not reference:
        raise InvalidArgumentError("refers_to cannot be empty.")
    check_formula(f"={reference}", workbook.sheetnames)
    names = _scope(workbook, sheet)
    replaced = name in names
    names[name] = DefinedName(name, attr_text=reference)
    return replaced


def delete_defined_name(workbook: Workbook, name: str, sheet: Worksheet | None) -> None:
    names = _scope(workbook, sheet)
    if name not in names:
        where = f"sheet {sheet.title!r}" if sheet else "the workbook"
        raise InvalidArgumentError(
            f"No name {name!r} is defined for {where}. "
            "describe_workbook lists defined names with their scope."
        )
    del names[name]


def _scope(workbook: Workbook, sheet: Worksheet | None) -> DefinedNameDict:
    return sheet.defined_names if sheet else workbook.defined_names
