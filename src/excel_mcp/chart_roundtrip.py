"""Keep what openpyxl drops or garbles when it reads a chart, which every later edit would save.

openpyxl rebuilds a chart from its file but forgets the chart style number, the rounded
corners flag (Excel draws a missing flag as rounded corners) and the plot area's fill. It
also gives every text paragraph an empty run, which comes back as the text "None" after the
next save, and it ignores which axes an area chart uses, so their settings are lost. Its data
label number formats are written without ``sourceLinked="0"``, which Excel reads as "use the
cell's format".
"""

# pyright: reportArgumentType=false, reportAttributeAccessIssue=false, reportOptionalMemberAccess=false
# openpyxl's stubs type its enum-like fields as Literals and its descriptors as plain attributes.
# ruff: noqa: N803  (the argument names are openpyxl's)

from openpyxl.chart import AreaChart, reader
from openpyxl.chart._chart import ChartBase
from openpyxl.chart.chartspace import ChartSpace
from openpyxl.chart.label import DataLabel, DataLabelList
from openpyxl.drawing import text
from openpyxl.reader import drawings
from openpyxl.xml.functions import Element

_read_chart = reader.read_chart
_init_paragraph = text.Paragraph.__init__
_init_area_chart = AreaChart.__init__
_labels_to_tree = DataLabelList.to_tree
_label_to_tree = DataLabel.to_tree


def _read_chart_completely(chartspace: ChartSpace) -> ChartBase:
    chart = _read_chart(chartspace)
    chart.style = chartspace.style
    chart.roundedCorners = chartspace.roundedCorners
    chart.plot_area.spPr = chartspace.chart.plotArea.spPr
    return chart


def _paragraph_without_empty_run(self, pPr=None, endParaRPr=None, r=None, br=None, fld=None):
    _init_paragraph(self, pPr, endParaRPr, [] if r is None else r, br, fld)


def _area_chart_with_axes(self, axId=None, **kw):
    _init_area_chart(self, **kw)
    self.axId = [] if axId is None else axId


def _unlink_number_format(tree: Element) -> Element:
    for node in tree.iter("numFmt"):
        node.set("sourceLinked", "0")
    return tree


def _labels_with_own_format(self, *args, **kw):
    return _unlink_number_format(_labels_to_tree(self, *args, **kw))


def _label_with_own_format(self, *args, **kw):
    return _unlink_number_format(_label_to_tree(self, *args, **kw))


def install() -> None:
    """Make openpyxl read charts completely. Safe to call again."""
    drawings.read_chart = _read_chart_completely
    text.Paragraph.__init__ = _paragraph_without_empty_run
    AreaChart.__init__ = _area_chart_with_axes
    DataLabelList.to_tree = _labels_with_own_format
    DataLabel.to_tree = _label_with_own_format
