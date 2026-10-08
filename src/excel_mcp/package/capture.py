"""Reading from the original file what openpyxl is going to drop."""

import zipfile
from pathlib import Path

from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from excel_mcp.errors import WorkbookError
from excel_mcp.package import properties
from excel_mcp.package.capture_sheet import capture_sheet, entries, part_extension
from excel_mcp.package.model import PackageState, WorkbookPackage, set_state
from excel_mcp.package.opc import REL_BASE
from excel_mcp.package.reader import Reader
from excel_mcp.package.scan import Scan, rel_id, scan, tags

_PACKAGE_REL = "http://schemas.openxmlformats.org/package/2006/relationships/"

WORKBOOK_ELEMENTS = frozenset(
    [
        "fileSharing",
        "functionGroups",
        "oleSize",
        "customWorkbookViews",
        "smartTagPr",
        "smartTagTypes",
        "webPublishing",
        "fileRecoveryPr",
        "webPublishObjects",
    ]
)
# Relationships openpyxl reads and writes again itself (or deliberately drops).
_WORKBOOK_RELS = {
    REL_BASE + name
    for name in [
        "worksheet",
        "chartsheet",
        "styles",
        "theme",
        "sharedStrings",
        "externalLink",
        "pivotCacheDefinition",
        "calcChain",
    ]
} | {"http://schemas.microsoft.com/office/2006/relationships/vbaProject"}
# A thumbnail shows the file as it was, and a digital signature is void once it is edited.
_ROOT_RELS = {
    REL_BASE + name for name in ["officeDocument", "extended-properties", "custom-properties"]
} | {
    _PACKAGE_REL + name
    for name in ["metadata/core-properties", "metadata/thumbnail", "digital-signature/origin"]
}


def capture(path: Path, workbook: Workbook, budget: int) -> None:
    """Remember what is in the file at ``path`` and that openpyxl dropped from ``workbook``.

    ``budget`` is how many bytes of such content may be kept in memory.
    """
    try:
        with zipfile.ZipFile(path) as archive:
            _capture(Reader(archive, budget), workbook)
    except (KeyError, ValueError, IndexError, UnicodeDecodeError, zipfile.BadZipFile) as error:
        raise WorkbookError(
            f"The package of {path.name} is damaged, so it cannot be edited safely ({error!r})."
        ) from None


def _capture(reader: Reader, workbook: Workbook) -> None:
    main = next(r.target for r in reader.rels("") if r.type == REL_BASE + "officeDocument")
    document = scan(reader.archive.read(main))
    state = PackageState(workbook=_workbook(reader, main, document))
    state.workbook.root_links = reader.links("", lambda rel: rel.type not in _ROOT_RELS)
    targets = {rel.id: rel.target for rel in reader.rels(main)}
    for values in tags(document, "sheet"):
        target = targets.get(rel_id(values, document))
        name = values["name"]
        if target is None or name not in workbook.sheetnames:
            continue
        sheet = workbook[name]
        state.sheet_ids[sheet] = int(values["sheetId"])
        state.sheet_names[sheet] = name
        if isinstance(sheet, Worksheet) and target in reader.names:
            state.sheets[sheet] = capture_sheet(reader, target, sheet, workbook.vba_archive is None)
    for sheet in workbook.worksheets:
        if isinstance(sheet, Worksheet):
            state.table_ids.update({table: table.id for table in sheet.tables.values()})
    set_state(workbook, state)


def _workbook(reader: Reader, main: str, document: Scan) -> WorkbookPackage:
    package = WorkbookPackage(
        company=properties.company_in(reader.archive), namespaces=document.namespaces
    )
    for child in document.children:
        if child.local in WORKBOOK_ELEMENTS:
            package.elements.append((child.local, document.raw(child)))
        elif child.local == "extLst":
            package.extensions = entries(document.raw(child), document.namespaces)
    package.links = reader.links(main, lambda rel: rel.type not in _WORKBOOK_RELS)
    targets = {rel.id: rel.target for rel in reader.rels(main)}
    for values in tags(document, "pivotCache"):
        target = targets.get(rel_id(values, document))
        if target in reader.names and (extension := part_extension(reader.archive.read(target))):
            package.cache_extensions[values["cacheId"]] = extension
    return package
