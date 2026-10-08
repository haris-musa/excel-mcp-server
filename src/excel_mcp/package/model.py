"""What the package layer remembers about a workbook between loading and saving it."""

from dataclasses import dataclass, field, fields
from typing import Any

from openpyxl import Workbook
from openpyxl.cell.cell import Cell
from openpyxl.chartsheet import Chartsheet
from openpyxl.worksheet.table import Table
from openpyxl.worksheet.worksheet import Worksheet

Sheet = Worksheet | Chartsheet


@dataclass(eq=False)
class Part:
    """A file inside the package. Its ``name`` is made unique when the package is written."""

    name: str
    content_type: str
    data: bytes
    links: list["Link"] = field(default_factory=list)
    by_default: bool = False


@dataclass(eq=False)
class Link:
    """A relationship from a part to another part (``Part``) or to a URL (``str``).

    ``id`` is what the owner's XML says (``r:id="rId3"``). On saving, it is replaced by a
    free id in the XML that is saved, so ids made up for new content only need to be unique.
    """

    id: str
    type: str
    target: "Part | str"


@dataclass(eq=False)
class CellMark:
    """Attributes of a cell that openpyxl does not model.

    ``cm`` points into ``xl/metadata.xml`` (dynamic arrays), ``vm`` too (linked data types,
    images in cells). ``dynamic`` asks for the dynamic array metadata, whatever its index.
    ``value`` is what the cell held when the mark was made: a ``vm`` only describes that
    value. ``cached`` is the stored result of an array formula as ``(type, text)``, and
    ``spill`` the cells it spilled into with the values they held.
    """

    cell: Cell
    cm: str | None = None
    vm: str | None = None
    dynamic: bool = False
    value: Any = None
    cached: tuple[str, str] | None = None
    spill: list[tuple[Cell, Any]] = field(default_factory=list)


@dataclass(eq=False)
class SheetPackage:
    """Content of one sheet that openpyxl drops.

    XML is carried as strings, exactly as Excel wrote it. Strings made by new code must
    declare the namespaces they use on their own element.

    - ``elements``: top-level elements of the sheet XML as (local name, XML) in file order.
    - ``extensions``: the ``<ext>`` entries of the sheet's ``<extLst>`` by uri.
    - ``anchors``: anchors of the sheet's drawing that openpyxl has no model for (slicers,
      new chart types, shapes), with the ``drawing_links`` they refer to.
    - ``links``: relationships of the sheet that the elements and extensions refer to,
      or that mean something by themselves (a threaded comments part).
    - ``vml``: the shapes of the sheet's VML drawing that are not notes (form controls, OLE
      objects), with the ``vml_links`` they refer to.
    - ``rule_extensions``: the ``<extLst>`` of conditional formatting rules by (range, priority).
    - ``pivot_extensions``: the ``<extLst>`` that closes each PivotTable, by name.
    - ``pivot_fields``: the ``<extLst>`` of a data field or page field of a PivotTable, by
      (pivot name, ``"dataField"`` or ``"pageField"``, position among those fields).
    - ``namespaces``: prefixes declared on the root of the sheet XML the strings come from.
    """

    elements: list[tuple[str, str]] = field(default_factory=list)
    extensions: dict[str, str] = field(default_factory=dict)
    anchors: list[str] = field(default_factory=list)
    drawing_links: list[Link] = field(default_factory=list)
    vml: list[str] = field(default_factory=list)
    vml_links: list[Link] = field(default_factory=list)
    links: list[Link] = field(default_factory=list)
    rule_extensions: dict[tuple[str, str], str] = field(default_factory=dict)
    pivot_extensions: dict[str, str] = field(default_factory=dict)
    pivot_fields: dict[tuple[str, str, int], str] = field(default_factory=dict)
    marks: list[CellMark] = field(default_factory=list)
    namespaces: dict[str, str] = field(default_factory=dict)
    drawing_namespaces: dict[str, str] = field(default_factory=dict)
    vml_namespaces: dict[str, str] = field(default_factory=dict)

    def is_empty(self) -> bool:
        return not any(
            getattr(self, f.name) for f in fields(self) if not f.name.endswith("namespaces")
        )

    def set_element(self, name: str, xml: str) -> None:
        """Make ``xml`` the only top-level element of the sheet with this local name."""
        self.elements = [(n, x) for n, x in self.elements if n != name] + [(name, xml)]

    def set_extension(self, uri: str, xml: str) -> None:
        """Add or replace the ``<ext>`` entry of the sheet's ``<extLst>`` with this uri."""
        self.extensions[uri] = xml

    def mark(self, mark: CellMark) -> None:
        """Attach ``mark`` to its cell in place of the cell's earlier one."""
        self.marks = [m for m in self.marks if m.cell is not mark.cell] + [mark]


@dataclass(eq=False)
class WorkbookPackage:
    """Content of the workbook that openpyxl drops: ``xl/workbook.xml`` and the package root.

    Like `SheetPackage`, with ``cache_extensions`` for the ``<extLst>`` of PivotCaches by
    cache id, ``root_links`` for the relationships of the package itself and ``company`` for
    the Company document property.
    """

    elements: list[tuple[str, str]] = field(default_factory=list)
    extensions: dict[str, str] = field(default_factory=dict)
    links: list[Link] = field(default_factory=list)
    root_links: list[Link] = field(default_factory=list)
    cache_extensions: dict[str, str] = field(default_factory=dict)
    company: str = ""
    namespaces: dict[str, str] = field(default_factory=dict)

    def is_empty(self) -> bool:
        return not any(
            getattr(self, f.name) for f in fields(self) if not f.name.endswith("namespaces")
        )

    def set_extension(self, uri: str, xml: str) -> None:
        """Add or replace the ``<ext>`` entry of the workbook's ``<extLst>`` with this uri."""
        self.extensions[uri] = xml


@dataclass(eq=False)
class PackageState:
    """Everything to put back when a workbook is saved: see the `excel_mcp.package` docstring."""

    workbook: WorkbookPackage = field(default_factory=WorkbookPackage)
    sheets: dict[Sheet, SheetPackage] = field(default_factory=dict)
    sheet_ids: dict[Sheet, int] = field(default_factory=dict)
    sheet_names: dict[Sheet, str] = field(default_factory=dict)
    table_ids: dict[Table, int] = field(default_factory=dict)

    def sheet(self, sheet: Sheet) -> SheetPackage:
        return self.sheets.setdefault(sheet, SheetPackage())

    def is_empty(self) -> bool:
        return self.workbook.is_empty() and all(p.is_empty() for p in self.sheets.values())


_KEY = "excel_mcp_package"


def workbook_sheets(workbook: Workbook) -> list[Sheet]:
    """All sheets of a workbook in order, chart sheets included."""
    return workbook._sheets  # pyright: ignore[reportAttributeAccessIssue]


def state_of(workbook: Workbook) -> PackageState:
    """The package state of a workbook, empty for one that was not loaded from a file."""
    return workbook.__dict__.setdefault(_KEY, PackageState())


def set_state(workbook: Workbook, state: PackageState) -> None:
    workbook.__dict__[_KEY] = state
