"""The XML of the Excel 2016 chart types (waterfall, histogram, ...), written as Excel writes it.

Excel keeps the ranges of these charts in hidden defined names (``_xlchart.v1.0``), which the
chart XML refers to; `NameBook` makes them.
"""

import re
import uuid
from dataclasses import dataclass
from xml.sax.saxutils import escape, quoteattr

from openpyxl import Workbook
from openpyxl.workbook.defined_name import DefinedName

from excel_mcp.operations.charts_data import Plot, split_sheet
from excel_mcp.operations.charts_options import ChartOptions, DataLabels
from excel_mcp.package.opc import XML_DECLARATION

_NAMESPACES = (
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
    'xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"'
)
_POSITIONS = {
    "center": "ctr", "inside_end": "inEnd", "inside_base": "inBase", "outside_end": "outEnd"
}  # fmt: skip
_LEGEND = {"right": "r", "left": "l", "top": "t", "bottom": "b"}


@dataclass(frozen=True)
class Kind:
    layout: str
    value_type: str
    gap_width: str
    value_axis: bool
    version: int


KINDS = {
    "waterfall": Kind("waterfall", "val", "0.5", True, 1),
    "histogram": Kind("clusteredColumn", "val", "0", True, 1),
    "pareto": Kind("clusteredColumn", "val", "0", True, 1),
    "box_whisker": Kind("boxWhisker", "val", "1", True, 1),
    "treemap": Kind("treemap", "size", "", False, 1),
    "sunburst": Kind("sunburst", "size", "", False, 1),
    "funnel": Kind("funnel", "val", "0.0599999987", False, 2),
}


class NameBook:
    """Makes the hidden names a chart's data lives under, numbered as Excel numbers them."""

    def __init__(self, workbook: Workbook) -> None:
        self.workbook = workbook
        self.numbers = [
            int(name.rpartition(".")[2])
            for name in workbook.defined_names
            if name.startswith("_xlchart.v") and name.rpartition(".")[2].isdigit()
        ]
        self.made: dict[tuple[str, int], str] = {}

    def of(self, reference: str, version: int) -> str:
        if (reference, version) not in self.made:
            number = max(self.numbers, default=-1) + 1
            self.numbers.append(number)
            name = f"_xlchart.v{version}.{number}"
            reference_text = _unquoted(reference)
            self.workbook.defined_names[name] = DefinedName(
                name, attr_text=reference_text, hidden=True
            )
            self.made[reference, version] = name
        return self.made[reference, version]


def _unquoted(reference: str) -> str:
    """The reference with its sheet name quoted only where Excel quotes it."""
    sheet, mark, cells = reference.rpartition("!")
    plain = sheet.removeprefix("'").removesuffix("'")
    if re.fullmatch(r"[A-Za-z_][A-Za-z0-9_.]*", plain) and not re.fullmatch(
        r"[A-Za-z]{1,3}\d+", plain
    ):
        return plain + mark + cells
    return reference


def chart_xml(workbook: Workbook, chart_type: str, plots: list[Plot], options: ChartOptions) -> str:
    kind = KINDS[chart_type]
    book = NameBook(workbook)
    categories = book.of(plots[0].categories, kind.version) if plots[0].categories else None
    data, series = [], []
    tail = str(uuid.uuid4()).upper()[8:]
    for number, plot in enumerate(plots):
        name = _name(workbook, book, plot, kind.version)
        values = book.of(plot.values, kind.version)
        dims = f'<cx:strDim type="cat"><cx:f>{categories}</cx:f></cx:strDim>' if categories else ""
        dims += f'<cx:numDim type="{kind.value_type}"><cx:f>{values}</cx:f></cx:numDim>'
        data.append(f'<cx:data id="{number}">{dims}</cx:data>')
        labels = _labels(plot.spec.data_labels or options.data_labels)
        parts = [
            name, labels, f'<cx:dataId val="{number}"/>', _layout(chart_type, options),
            '<cx:axisId val="1"/>' if chart_type == "pareto" else "",
        ]  # fmt: skip
        series.append(_series(kind.layout, f"{{{number + 1:08X}{tail}}}", "".join(parts)))
    if chart_type == "pareto":
        line_id = f"{{{str(uuid.uuid4()).upper()}}}"
        series.append(
            f'<cx:series layoutId="paretoLine" ownerIdx="0" uniqueId="{line_id}">'
            '<cx:axisId val="2"/></cx:series>'
        )
    body = (
        f"{_title(options.title)}<cx:plotArea><cx:plotAreaRegion>{''.join(series)}"
        f"</cx:plotAreaRegion>{_axes(chart_type, options)}</cx:plotArea>{_legend(options)}"
    )
    return (
        f"{XML_DECLARATION}<cx:chartSpace {_NAMESPACES}><cx:chartData>{''.join(data)}"
        f"</cx:chartData><cx:chart>{body}</cx:chart></cx:chartSpace>"
    )


def _series(layout: str, unique_id: str, body: str) -> str:
    return f'<cx:series layoutId="{layout}" uniqueId="{unique_id}">{body}</cx:series>'


def _name(workbook: Workbook, book: NameBook, plot: Plot, version: int) -> str:
    label = plot.name
    if label is None:
        return ""
    if label.v is not None:
        return f"<cx:tx><cx:txData><cx:v>{escape(label.v)}</cx:v></cx:txData></cx:tx>"
    assert label.strRef is not None
    reference = label.strRef.f or ""
    title, cell = split_sheet(workbook, "", reference)
    text = workbook[title][cell.replace("$", "")].value
    shown = "" if text is None else escape(str(text))
    name = book.of(reference, version)
    return f"<cx:tx><cx:txData><cx:f>{name}</cx:f><cx:v>{shown}</cx:v></cx:txData></cx:tx>"


def _title(text: str | None) -> str:
    if text is None:
        return ""
    return (
        '<cx:title pos="t" align="ctr" overlay="0"><cx:tx><cx:txData>'
        f"<cx:v>{escape(text)}</cx:v></cx:txData></cx:tx></cx:title>"
    )


def _legend(options: ChartOptions) -> str:
    if options.legend == "none":
        return ""
    return f'<cx:legend pos="{_LEGEND[options.legend]}" align="ctr" overlay="0"/>'


def _labels(labels: DataLabels | None) -> str:
    if labels is None:
        return ""
    shown = {
        attribute: int(name in labels.show)
        for attribute, name in (
            ("seriesName", "series"),
            ("categoryName", "category"),
            ("value", "value"),
        )
    }
    position = f" pos={quoteattr(_POSITIONS[labels.position])}" if labels.position else ""
    number = (
        f'<cx:numFmt formatCode={quoteattr(labels.number_format)} sourceLinked="0"/>'
        if labels.number_format
        else ""
    )
    visibility = " ".join(f'{key}="{value}"' for key, value in shown.items())
    return f"<cx:dataLabels{position}>{number}<cx:visibility {visibility}/></cx:dataLabels>"


def _layout(chart_type: str, options: ChartOptions) -> str:
    if chart_type == "waterfall":
        lines = "" if options.connector_lines else '<cx:visibility connectorLines="0"/>'
        totals = "".join(f'<cx:idx val="{n - 1}"/>' for n in sorted(set(options.totals)))
        return f"<cx:layoutPr>{lines}<cx:subtotals>{totals}</cx:subtotals></cx:layoutPr>"
    if chart_type == "histogram":
        return f"<cx:layoutPr>{_binning(options)}</cx:layoutPr>"
    if chart_type == "pareto":
        return "<cx:layoutPr><cx:aggregation/></cx:layoutPr>"
    if chart_type == "box_whisker":
        return f"<cx:layoutPr>{_box(options)}</cx:layoutPr>"
    if chart_type == "treemap":
        label = options.parent_labels
        found = "" if label == "none" else f'<cx:parentLabelLayout val="{label}"/>'
        return f"<cx:layoutPr>{found}</cx:layoutPr>"
    return ""


def _binning(options: ChartOptions) -> str:
    bins = options.bins
    edges = "".join(
        f' {name}="{_number(value)}"'
        for name, value in (("underflow", bins.underflow), ("overflow", bins.overflow))
        if value is not None
    )
    size = ""
    if bins.width is not None:
        size = f'<cx:binSize val="{_number(bins.width)}"/>'
    elif bins.count is not None:
        size = f'<cx:binCount val="{bins.count}"/>'
    return f'<cx:binning intervalClosed="r"{edges}>{size}</cx:binning>'


def _box(options: ChartOptions) -> str:
    box = options.box
    shown = ""
    if (box.mean_marker, box.mean_line, box.inner_points, box.outliers) != (
        True,
        False,
        False,
        True,
    ):
        shown = (
            f'<cx:visibility meanLine="{int(box.mean_line)}" meanMarker="{int(box.mean_marker)}" '
            f'nonoutliers="{int(box.inner_points)}" outliers="{int(box.outliers)}"/>'
        )
    return f'{shown}<cx:statistics quartileMethod="{box.quartiles}"/>'


def _axes(chart_type: str, options: ChartOptions) -> str:
    kind = KINDS[chart_type]
    if not kind.gap_width:
        return ""
    category = (
        f'<cx:axis id="0"><cx:catScaling gapWidth="{kind.gap_width}"/>'
        f"{_axis_title(options.x_axis.title)}<cx:tickLabels/></cx:axis>"
    )
    if not kind.value_axis:
        return category
    axis = options.y_axis
    limits = "".join(
        f' {name}="{_number(value)}"'
        for name, value in (("max", axis.max), ("min", axis.min), ("majorUnit", axis.major_unit))
        if value is not None
    )
    grid = "" if axis.major_gridlines is False else "<cx:majorGridlines/>"
    number = (
        f'<cx:numFmt formatCode={quoteattr(axis.number_format)} sourceLinked="0"/>'
        if axis.number_format
        else ""
    )
    value = (
        f'<cx:axis id="1"><cx:valScaling{limits}/>{_axis_title(axis.title)}{grid}'
        f"<cx:tickLabels/>{number}</cx:axis>"
    )
    percent = (
        '<cx:axis id="2"><cx:valScaling max="1" min="0"/><cx:units unit="percentage"/>'
        "<cx:tickLabels/></cx:axis>"
        if chart_type == "pareto"
        else ""
    )
    return category + value + percent


def _axis_title(text: str | None) -> str:
    if text is None:
        return ""
    return f"<cx:title><cx:tx><cx:txData><cx:v>{escape(text)}</cx:v></cx:txData></cx:tx></cx:title>"


def _number(value: float) -> str:
    return str(int(value)) if value == int(value) else repr(value)
