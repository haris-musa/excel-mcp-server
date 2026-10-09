"""Keep what openpyxl drops or garbles when it reads a chart, which every later edit would save.

openpyxl rebuilds a chart from its file but forgets the chart style number, the rounded
corners flag (Excel draws a missing flag as rounded corners) and the plot area's fill. It
also gives every text paragraph an empty run, which comes back as the text "None" after the
next save, and it ignores which axes an area chart uses, so their settings are lost. Its data
label number formats are written without ``sourceLinked="0"``, which Excel reads as "use the
cell's format". It writes every chart as "Chart <n>" and every picture
as "Image <n>", so charts and pictures are written with the names they have instead.
"""

# pyright: reportArgumentType=false, reportAttributeAccessIssue=false, reportOptionalMemberAccess=false
# openpyxl's stubs type its enum-like fields as Literals and its descriptors as plain attributes.
# ruff: noqa: N803  (the argument names are openpyxl's)

from openpyxl.chart import AreaChart, reader
from openpyxl.chart._chart import ChartBase
from openpyxl.chart.chartspace import ChartSpace
from openpyxl.chart.label import DataLabel, DataLabelList
from openpyxl.drawing import text
from openpyxl.drawing.spreadsheet_drawing import SpreadsheetDrawing
from openpyxl.reader import drawings
from openpyxl.xml.functions import Element

from excel_mcp.package.shape_names import name_of

_read_chart = reader.read_chart
_init_paragraph = text.Paragraph.__init__
_init_area_chart = AreaChart.__init__
_labels_to_tree = DataLabelList.to_tree
_label_to_tree = DataLabel.to_tree
_write_drawing = SpreadsheetDrawing._write
_chart_frame = SpreadsheetDrawing._chart_frame
_picture_frame = SpreadsheetDrawing._picture_frame


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


def _write_named_drawing(self):
    self.shape_names = [name_of(shape) for shape in self.charts + self.images]
    for shape in self.images:
        anchor = shape.anchor
        if not isinstance(anchor, str) and anchor.pic is not None and name_of(shape):
            anchor.pic.nvPicPr.cNvPr.name = name_of(shape)
    return _write_drawing(self)


def _named_chart_frame(self, idx):
    frame = _chart_frame(self, idx)
    frame.nvGraphicFramePr.cNvPr.name = (
        self.shape_names[idx - 1] or frame.nvGraphicFramePr.cNvPr.name
    )
    return frame


def _named_picture_frame(self, idx):
    frame = _picture_frame(self, idx)
    frame.nvPicPr.cNvPr.name = self.shape_names[idx - 1] or frame.nvPicPr.cNvPr.name
    return frame


def install() -> None:
    """Make openpyxl read charts completely. Safe to call again."""
    drawings.read_chart = _read_chart_completely
    text.Paragraph.__init__ = _paragraph_without_empty_run
    AreaChart.__init__ = _area_chart_with_axes
    DataLabelList.to_tree = _labels_with_own_format
    DataLabel.to_tree = _label_with_own_format
    SpreadsheetDrawing._write = _write_named_drawing
    SpreadsheetDrawing._chart_frame = _named_chart_frame
    SpreadsheetDrawing._picture_frame = _named_picture_frame
