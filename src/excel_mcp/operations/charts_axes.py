"""Axes: scale, titles, gridlines, labels, and the second value axis."""

# pyright: reportArgumentType=false, reportAttributeAccessIssue=false, reportOptionalMemberAccess=false
# openpyxl's stubs type its enum-like fields as Literals and its descriptors as plain attributes.

from openpyxl.chart.axis import NumericAxis, TextAxis
from openpyxl.chart.data_source import NumFmt
from openpyxl.chart.shapes import GraphicalProperties

from excel_mcp.operations import charts_look as look
from excel_mcp.operations.charts_options import Axis

SECONDARY_X, SECONDARY_Y = 500, 200
_VERTICAL_TITLE = -5400000
_LABEL_POSITIONS = {"next_to_axis": "nextTo", "low": "low", "high": "high"}


def style_axis(
    axis: NumericAxis | TextAxis,
    spec: Axis,
    *,
    vertical: bool,
    default_gridlines: bool,
    line: tuple[int, int] | None,
) -> None:
    """Apply `spec` to one axis, drawn as current Excel draws it.

    `line` is the (lumMod, lumOff) of the axis line, or None for no line. `vertical` tells
    which way the axis title reads.
    """
    # openpyxl marks axes as deleted by default, which hides them in current Excel.
    axis.delete = False
    if spec.title:
        size = 1000
        axis.title = look.title(spec.title, size, _VERTICAL_TITLE if vertical else None)
    gridlines = default_gridlines if spec.major_gridlines is None else spec.major_gridlines
    axis.majorGridlines = look.gridlines() if gridlines else None
    axis.minorGridlines = look.gridlines(minor=True) if spec.minor_gridlines else None
    axis.tickLblPos = _LABEL_POSITIONS[spec.labels]
    axis.crosses = "autoZero"
    axis.majorTickMark = "none"
    axis.minorTickMark = "none"
    axis.numFmt = NumFmt(
        formatCode=spec.number_format or "General", sourceLinked=not spec.number_format
    )
    scaling = axis.scaling
    scaling.min, scaling.max = spec.min, spec.max
    scaling.logBase = 10 if spec.log else None
    scaling.orientation = "maxMin" if spec.reverse else "minMax"
    if isinstance(axis, NumericAxis):
        axis.majorUnit = spec.major_unit
    else:
        axis.auto = True
        axis.lblAlgn = "ctr"
        axis.lblOffset = 100
        axis.noMultiLvlLbl = False
    axis.txPr = look.text_properties(900, rotation=-60000000)
    axis.spPr = (
        GraphicalProperties(noFill=True, ln=look.thin_line(*line)) if line else look.no_fill()
    )


def secondary_axes(group, spec: Axis, *, scatter: bool, cross: str) -> None:
    """Give a plot group its own pair of axes, with the value axis on the right.

    The second category (or x) axis is only there for the pairing, so it is hidden.
    """
    group.x_axis.axId, group.y_axis.axId = SECONDARY_X, SECONDARY_Y
    group.x_axis.crossAx, group.y_axis.crossAx = SECONDARY_Y, SECONDARY_X
    style_axis(
        group.y_axis, spec, vertical=True, default_gridlines=False,
        line=(25000, 75000) if scatter else None,
    )  # fmt: skip
    group.y_axis.axPos = "r"
    group.y_axis.crossBetween = cross
    group.y_axis.crosses = "max"
    style_axis(group.x_axis, Axis(), vertical=False, default_gridlines=False, line=None)
    group.x_axis.delete = True
    group.x_axis.axPos = "b"
    if scatter:
        group.x_axis.crossBetween = cross
