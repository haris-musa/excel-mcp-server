"""Keeping the VBA project of a macro-enabled workbook inside an openpyxl workbook.

openpyxl keeps the macro parts of an .xlsm file in ``workbook.vba_archive`` and copies
them into the file when it saves. Changing the project means replacing that archive.
"""

import io
import zipfile

from openpyxl import Workbook
from openpyxl.chartsheet import Chartsheet
from openpyxl.xml.functions import Element, SubElement, fromstring, tostring

from excel_mcp.ovba_write import Project

VBA_PART = "xl/vbaProject.bin"
_CONTENT_TYPES_PART = "[Content_Types].xml"
_ROOT_RELS_PART = "_rels/.rels"
_VBA_CONTENT_TYPE = "application/vnd.ms-office.vbaProject"
_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
_EMPTY_RELS = (
    b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>'
)

_WORKBOOK_BASE = "{00020819-0000-0000-C000-000000000046}"
_WORKSHEET_BASE = "{00020820-0000-0000-C000-000000000046}"
_CHARTSHEET_BASE = "{00020821-0000-0000-C000-000000000046}"


def read_project(workbook: Workbook) -> Project | None:
    archive = workbook.vba_archive
    if archive is None or VBA_PART not in archive.namelist():
        return None
    return Project.from_bytes(archive.read(VBA_PART))


def new_project(workbook: Workbook) -> Project:
    """An empty project with a document module for the workbook and each of its sheets.

    Excel pairs those modules with the workbook and sheets by their code names, so
    names are given to the ones that have none yet.
    """
    workbook.code_name = workbook.code_name or "ThisWorkbook"
    documents = [(workbook.code_name, _WORKBOOK_BASE)]
    sheets = [workbook[name] for name in workbook.sheetnames]
    taken = {workbook.code_name.casefold()}
    taken |= {
        name.casefold() for sheet in sheets if (name := sheet.sheet_properties.codeName) is not None
    }
    for sheet in sheets:
        is_chartsheet = isinstance(sheet, Chartsheet)
        if sheet.sheet_properties.codeName is None:
            prefix = "Chart" if is_chartsheet else "Sheet"
            sheet.sheet_properties.codeName = _free_name(prefix, taken)
        base = _CHARTSHEET_BASE if is_chartsheet else _WORKSHEET_BASE
        documents.append((sheet.sheet_properties.codeName, base))
    return Project.new(documents)


def store_project(workbook: Workbook, project: Project) -> None:
    """Make the workbook save with this project, adding the macro parts it lacks."""
    old = workbook.vba_archive
    kept: dict[str, bytes] = {}
    if old is not None:
        kept = {name: old.read(name) for name in old.namelist()}
        old.close()
    kept[VBA_PART] = project.to_bytes()
    kept[_CONTENT_TYPES_PART] = _with_vba_override(kept.get(_CONTENT_TYPES_PART))
    kept.setdefault(_ROOT_RELS_PART, _EMPTY_RELS)
    archive = zipfile.ZipFile(io.BytesIO(), "a", zipfile.ZIP_DEFLATED)
    for name, content in kept.items():
        archive.writestr(name, content)
    workbook.vba_archive = archive


def _free_name(prefix: str, taken: set[str]) -> str:
    number = 1
    while f"{prefix}{number}".casefold() in taken:
        number += 1
    taken.add(f"{prefix}{number}".casefold())
    return f"{prefix}{number}"


def _with_vba_override(content_types: bytes | None) -> bytes:
    types = Element(f"{{{_TYPES_NS}}}Types") if content_types is None else fromstring(content_types)
    if not any(override.get("PartName") == f"/{VBA_PART}" for override in types):
        SubElement(
            types,
            f"{{{_TYPES_NS}}}Override",
            PartName=f"/{VBA_PART}",
            ContentType=_VBA_CONTENT_TYPE,
        )
    # openpyxl's type stub says str, but it returns bytes.
    return tostring(types)  # pyright: ignore[reportReturnType]
