"""Read-only inspection of the VBA macros in .xlsm and .xltm workbooks."""

import zipfile
from pathlib import Path

from pydantic import BaseModel

from excel_mcp import ovba
from excel_mcp.errors import InvalidArgumentError, WorkbookError
from excel_mcp.ovba import ModuleKind

VBA_PROJECT_PART = "xl/vbaProject.bin"


class VbaModule(BaseModel):
    name: str
    kind: ModuleKind
    line_count: int
    code: str


class VbaProject(BaseModel):
    modules: list[VbaModule]
    truncated: bool


def has_vba(path: Path) -> bool:
    with zipfile.ZipFile(path) as archive:
        return VBA_PROJECT_PART in archive.namelist()


def read_vba(path: Path, module: str | None, max_chars: int) -> VbaProject:
    raw_modules = ovba.read_project(_project_bytes(path)).modules
    modules = [_module(raw) for raw in raw_modules]
    if module is not None:
        modules = [found for found in modules if found.name.casefold() == module.casefold()]
        if not modules:
            names = ", ".join(repr(raw.name) for raw in raw_modules)
            raise InvalidArgumentError(f"Module {module!r} not found. Modules: {names}.")
    return _within_budget(modules, max_chars)


def editor_text(source: str) -> str:
    """The code as the VBA editor shows it: without hidden attributes, with \\n line ends."""
    lines = source.replace("\r\n", "\n").split("\n")
    return "\n".join(line for line in lines if not line.startswith("Attribute ")).strip("\n")


def _project_bytes(path: Path) -> bytes:
    if not zipfile.is_zipfile(path):
        raise WorkbookError("The file is not an Excel workbook.")
    with zipfile.ZipFile(path) as archive:
        if VBA_PROJECT_PART not in archive.namelist():
            raise InvalidArgumentError("This workbook contains no VBA macros.")
        info = archive.getinfo(VBA_PROJECT_PART)
        if info.file_size > ovba.MAX_DECOMPRESSED_BYTES:
            raise WorkbookError("The VBA project is too large to read.")
        return archive.read(info)


def _module(raw: ovba.RawModule) -> VbaModule:
    code = editor_text(raw.source)
    return VbaModule(name=raw.name, kind=raw.kind, line_count=len(code.splitlines()), code=code)


def _within_budget(modules: list[VbaModule], max_chars: int) -> VbaProject:
    remaining = max_chars
    truncated = False
    limited: list[VbaModule] = []
    for module in modules:
        code = module.code[:remaining]
        truncated = truncated or len(code) < len(module.code)
        remaining -= len(code)
        limited.append(module.model_copy(update={"code": code}))
    return VbaProject(modules=limited, truncated=truncated)
