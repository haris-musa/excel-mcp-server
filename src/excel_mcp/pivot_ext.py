"""Keeping the extension elements of PivotTables.

Excel stores newer PivotTable settings (rank and percent-of-parent figures, hidden "Values"
rows) in ``extLst`` elements that openpyxl reads and writes back empty. This module makes the
``extLst`` of every pivot part keep its content, so editing a workbook does not drop it.
"""

import copy

from openpyxl.descriptors.excel import ExtensionList
from openpyxl.pivot import table
from openpyxl.xml.functions import Element, SubElement

MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
X14 = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
SHOW_VALUES_AS = "{E15A36E0-9728-4e99-A89B-3F7291B0FE68}"


class OpaqueExtensions(ExtensionList):
    """An ``extLst`` that keeps its child elements as they are."""

    def __init__(self, children=()) -> None:
        super().__init__()
        self.children = list(children)

    @classmethod
    def from_tree(cls, node):
        return cls(children=[copy.deepcopy(child) for child in node])

    def to_tree(self, tagname=None, idx=None, namespace=None):
        element = Element("extLst")
        element.extend(copy.deepcopy(child) for child in self.children)
        return element


def show_values_as(kind: str) -> OpaqueExtensions:
    """The extension that makes a data field show rank or percent-of-parent figures."""
    ext = Element(f"{{{MAIN}}}ext", {"uri": SHOW_VALUES_AS})
    SubElement(ext, f"{{{X14}}}dataField", {"pivotShowAs": kind})
    return OpaqueExtensions([ext])


def keep_extensions() -> None:
    """Make PivotTable parts keep their extension elements. Safe to call again."""
    for owner in (table.TableDefinition, table.DataField, table.PageField):
        descriptor = owner.__dict__["extLst"]
        if descriptor.expected_type is not OpaqueExtensions:
            descriptor.expected_type = OpaqueExtensions
            owner.__elements__ = (*owner.__elements__, "extLst")
