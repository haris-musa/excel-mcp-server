"""How many tables, charts, PivotTables, slicers and pictures each sheet of a file holds.

Read straight from the package, so that describe_workbook stays a streaming read.
"""

import zipfile
from pathlib import Path
from xml.etree import ElementTree

from pydantic import BaseModel

from excel_mcp.errors import WorkbookError
from excel_mcp.package.opc import REL_BASE, Rel, parse_rels, rels_name
from excel_mcp.package.scan import rel_id, scan, tags

_DRAWING = "{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}"
_GRAPHIC_DATA = "{http://schemas.openxmlformats.org/drawingml/2006/main}graphicData"
_CHARTS = (
    "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "http://schemas.microsoft.com/office/drawing/2014/chartex",
)
_SLICERS = (
    "http://schemas.microsoft.com/office/drawing/2010/slicer",
    "http://schemas.microsoft.com/office/drawing/2012/timeslicer",
)


class ObjectCounts(BaseModel):
    tables: int = 0
    charts: int = 0
    pivot_tables: int = 0
    slicers: int = 0
    images: int = 0


def count_objects(path: Path) -> dict[str, ObjectCounts]:
    """The objects of each worksheet by sheet name."""
    try:
        with zipfile.ZipFile(path) as archive:
            return _count(archive)
    except (KeyError, ValueError, ElementTree.ParseError, zipfile.BadZipFile) as error:
        raise WorkbookError(f"The package of {path.name} is damaged ({error!r}).") from None


def _count(archive: zipfile.ZipFile) -> dict[str, ObjectCounts]:
    names = set(archive.namelist())

    def rels(part: str) -> list[Rel]:
        name = rels_name(part)
        return parse_rels(part, archive.read(name)) if name in names else []

    main = next(r.target for r in rels("") if r.type == REL_BASE + "officeDocument")
    targets = {rel.id: rel for rel in rels(main)}
    document = scan(archive.read(main))
    counts = {}
    for values in tags(document, "sheet"):
        sheet = targets.get(rel_id(values, document))
        if sheet is not None and sheet.type == REL_BASE + "worksheet" and sheet.target in names:
            counts[values["name"]] = _sheet(archive, names, rels(sheet.target))
    return counts


def _sheet(archive: zipfile.ZipFile, names: set[str], rels: list[Rel]) -> ObjectCounts:
    counts = ObjectCounts()
    for rel in rels:
        if rel.type == REL_BASE + "table":
            counts.tables += 1
        elif rel.type == REL_BASE + "pivotTable":
            counts.pivot_tables += 1
        elif rel.type == REL_BASE + "drawing" and rel.target in names:
            _drawing(ElementTree.fromstring(archive.read(rel.target)), counts)
    return counts


def _drawing(root: ElementTree.Element, counts: ObjectCounts) -> None:
    for element in root.iter():
        if element.tag == f"{_DRAWING}pic":
            counts.images += 1
        elif element.tag == _GRAPHIC_DATA:
            uri = element.get("uri")
            counts.charts += uri in _CHARTS
            counts.slicers += uri in _SLICERS
