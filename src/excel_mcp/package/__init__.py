"""The package layer: keeps what openpyxl cannot model across an edit.

openpyxl loads the parts of a workbook it has a model for and writes only those back.
Everything else is gone after a load and a save: sparklines, slicers, newer chart types,
threaded comments, extensions of conditional formats, cell metadata and more. This layer
reads those parts from the original file when a workbook is opened for editing, and puts
them into the file openpyxl wrote, so an edit changes what it was asked to change.

Preserved content is carried as `Part` objects (files of the package) and as XML text,
exactly as Excel wrote it; nothing in it is parsed and written again.

Life cycle (`Workspace.edit` does all of it)::

    load_workbook(path)           # openpyxl
    capture(path, workbook, ...)  # remember what openpyxl dropped -> state_of(workbook)
    ... the tool edits the workbook, and may add to or change the state ...
    prepare(workbook)             # make the workbook consistent with the state
    workbook.save(buffer)         # openpyxl
    apply(workbook, buffer)       # put the state back into the file openpyxl wrote

Adding content, as a feature that writes sparklines, slicers or comments does::

    state = state_of(workbook)
    sheet = state.sheet(worksheet)   # a SheetPackage; state.workbook is the WorkbookPackage
    sheet.set_extension(uri, '<ext uri="...">...</ext>')    # entry of the sheet's <extLst>
    sheet.set_element("legacyDrawingHF", xml)               # element openpyxl has no model for
    sheet.anchors.append(xml)                               # shape in the sheet's drawing
    sheet.links.append(Link("rIdNew", rel_type, Part(name, content_type, data)))

- XML is a string with the namespaces it uses declared on its own element. Relationship ids
  inside it (``r:id="rIdNew"``) are replaced by free ones when the file is written, so any
  id that is unique among the ones you add will do. Elements are placed in schema order;
  ``<extLst>`` is built from the extension entries.
- A `Part` is a file: its name is made unique if taken, its content type is declared, and
  the parts it links to are written too. A part that is the same object is written once.
- `CellMark` sets the ``cm`` and ``vm`` attributes of a cell (dynamic arrays, linked data
  types); ``cm`` is written while the cell holds an array formula, ``vm`` while it holds the
  value it described.
- `package.references` moves preserved content when rows or columns are inserted or deleted
  (`rewrite_lines(workbook, LineEdit(...), formula)`) and when formulas change for other
  reasons, such as a sheet rename (`rewrite_formulas(workbook, formula)`). It covers
  sparklines, Excel 2010 conditional formats and validations, shape anchors (drawing, VML and
  form controls), control properties, newer chart data, ignored errors, protected ranges,
  cell watches and sort state, with Excel's own rules. The caller supplies ``formula(text,
  host)``, its rewriting of one formula; read the module docstring for the contract.
  `package.extensions.forget_sheets` handles deleted sheets.
- `package.guards` refuses edits that would leave content this server cannot store; call
  `check_sheet_removal` and `check_pivot_removal` before deleting.
- Content that refers to what an edit removed is cleaned up as Excel does it when the file
  is written (`package.consistency`): slicer caches no slicer uses, threaded comments
  without their note, sparklines that lose their data with a deleted sheet.
- Sheet and table numbers, which openpyxl renumbers, are remapped in the slicer and timeline
  caches that name them; parts you add may name sheets and tables by the numbers they had
  in the file.

The size of what is kept is limited by the server's file size limit.
"""

from openpyxl import Workbook

from excel_mcp.package import arrays, consistency, pivot_records
from excel_mcp.package.capture import capture
from excel_mcp.package.model import (
    CellMark,
    Link,
    PackageState,
    Part,
    SheetPackage,
    WorkbookPackage,
    state_of,
)
from excel_mcp.package.restore import apply


def prepare(workbook: Workbook) -> None:
    """Make the workbook consistent with its preserved content; call before saving it."""
    arrays.prepare(workbook)
    consistency.drop_orphaned_caches(workbook)
    pivot_records.prepare(workbook)


__all__ = [
    "CellMark",
    "Link",
    "PackageState",
    "Part",
    "SheetPackage",
    "WorkbookPackage",
    "apply",
    "capture",
    "prepare",
    "state_of",
]
