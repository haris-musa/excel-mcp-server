"""Writing VBA macro code into .xlsm and .xltm workbooks. The code is stored, never run."""

import re
from typing import Literal

from openpyxl import Workbook
from pydantic import BaseModel

from excel_mcp import macros
from excel_mcp.errors import InvalidArgumentError
from excel_mcp.ovba import ModuleKind
from excel_mcp.ovba_write import Module, Project

MAX_CODE_CHARS = 200_000

_MODULE_NAME = re.compile(r"[A-Za-z][A-Za-z0-9_]{0,30}")
_ATTRIBUTE_LINE = re.compile(r"^attribute\b", re.IGNORECASE | re.MULTILINE)


class VbaChange(BaseModel):
    module: str
    kind: ModuleKind
    action: Literal["created", "replaced", "deleted"]


def write_module(
    workbook: Workbook, name: str, code: str, kind: Literal["standard", "class"]
) -> VbaChange:
    if _ATTRIBUTE_LINE.search(code):
        raise InvalidArgumentError(
            "Code lines cannot start with 'Attribute': the server writes module attributes itself."
        )
    if "\0" in code:
        raise InvalidArgumentError("The code contains a NUL character.")
    project = _project(workbook)
    existing = project.find(name)
    if existing is None and not _MODULE_NAME.fullmatch(name):
        raise InvalidArgumentError(
            f"{name!r} is not a valid module name: use 1-31 letters, digits or underscores, "
            "starting with a letter."
        )
    if existing is not None and existing.kind == "form":
        raise InvalidArgumentError(f"{existing.name!r} is a form; forms cannot be changed.")
    created = project.set_code(name, code, kind)
    macros.store_project(workbook, project)
    module = existing or project.modules[-1]
    return VbaChange(
        module=module.name, kind=module.kind, action="created" if created else "replaced"
    )


def delete_module(workbook: Workbook, name: str) -> VbaChange:
    project = _project(workbook)
    module = _find(project, name)
    if module.kind not in ("standard", "class"):
        raise InvalidArgumentError(
            f"{module.name!r} is a {module.kind} module and cannot be deleted. "
            "Replace its code with an empty string to clear it."
        )
    project.remove(module)
    macros.store_project(workbook, project)
    return VbaChange(module=module.name, kind=module.kind, action="deleted")


def _project(workbook: Workbook) -> Project:
    if workbook.vba_archive is None:
        raise InvalidArgumentError(
            "Macros can only be stored in .xlsm or .xltm workbooks, not in .xlsx or .xltx files."
        )
    if any(part.startswith("xl/vbaProjectSignature") for part in workbook.vba_archive.namelist()):
        raise InvalidArgumentError(
            "This workbook's VBA project is digitally signed; changing it would invalidate "
            "the signature. Remove the signature in Excel (Tools > Digital Signature) or work "
            "on an unsigned copy."
        )
    return macros.read_project(workbook) or macros.new_project(workbook)


def _find(project: Project, name: str) -> Module:
    module = project.find(name)
    if module is None:
        names = ", ".join(repr(module.name) for module in project.modules)
        raise InvalidArgumentError(f"Module {name!r} not found. Modules: {names}.")
    return module
