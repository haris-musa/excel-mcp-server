"""Sparkline groups as Excel 2010 stores them in a sheet's extensions: writing and reading."""

from xml.etree import ElementTree
from xml.sax.saxutils import escape

from excel_mcp.operations.sparkline_style import (
    LINE_WEIGHT,
    AxisScale,
    SparklineColors,
    SparklineInfo,
    SparklineStyle,
)
from excel_mcp.operations.x14_xml import X14, XM, attributes, color, number_text

_TYPES = {"line": None, "column": "column", "win_loss": "stacked"}
_EMPTY_CELLS = {"gap": "gap", "zero": None, "connect": "span"}
_POINTS = ("markers", "high", "low", "first", "last", "negative")
_COLORS = ("series", "negative", "axis", "markers", "first", "last", "high", "low")
_SCALES = {"same": "group"}


def build_group(style: SparklineStyle, sparklines: list[tuple[str, str]]) -> str:
    """One ``<x14:sparklineGroup>`` for (cell, data formula) pairs, written as Excel writes it."""
    colors = "".join(
        color(f"color{name.capitalize()}", getattr(style.colors, name)) for name in _COLORS
    )
    dates = f"<xm:f>{escape(style.dates)}</xm:f>" if style.dates else ""
    items = "".join(
        f"<x14:sparkline><xm:f>{escape(data)}</xm:f><xm:sqref>{cell}</xm:sqref></x14:sparkline>"
        for cell, data in sparklines
    )
    return (
        f"<x14:sparklineGroup{_attributes(style)}>{colors}{dates}"
        f"<x14:sparklines>{items}</x14:sparklines></x14:sparklineGroup>"
    )


def _attributes(style: SparklineStyle) -> str:
    def flag(on: bool) -> str | None:
        return "1" if on else None

    def custom(scale: AxisScale) -> str | None:
        return None if isinstance(scale, str) else number_text(scale)

    def kind(scale: AxisScale) -> str | None:
        return _SCALES.get(scale) if isinstance(scale, str) else "custom"

    weight = number_text(style.line_weight) if style.line_weight != LINE_WEIGHT else None
    return attributes(
        [
            ("manualMax", custom(style.axis_max)),
            ("manualMin", custom(style.axis_min)),
            ("lineWeight", weight),
            ("type", _TYPES[style.type]),
            ("dateAxis", flag(style.dates is not None)),
            ("displayEmptyCellsAs", _EMPTY_CELLS[style.empty_cells]),
            *[(point, flag(point in style.show)) for point in _POINTS],
            ("displayXAxis", flag(style.show_axis)),
            ("displayHidden", flag(style.hidden)),
            ("minAxisType", kind(style.axis_min)),
            ("maxAxisType", kind(style.axis_max)),
            ("rightToLeft", flag(style.right_to_left)),
        ]
    )


def read_groups(xml: str) -> list[SparklineInfo]:
    """The sparkline groups of the sheet's sparkline extension."""
    root = ElementTree.fromstring(xml)
    return [_read_group(group) for group in root.iter(f"{{{X14}}}sparklineGroup")]


def _read_group(group: ElementTree.Element) -> SparklineInfo:
    def flag(name: str) -> bool:
        return group.get(name) in ("1", "true")

    def scale(side: str) -> AxisScale:
        kind = group.get(f"{side}AxisType", "individual")
        if kind == "custom":
            return float(group.get(f"manual{side.capitalize()}", "0"))
        return "same" if kind == "group" else "individual"

    types = {value: key for key, value in _TYPES.items() if value}
    empty = {value: key for key, value in _EMPTY_CELLS.items() if value}
    dates = group.find(f"{{{XM}}}f")
    sparklines = {
        item.findtext(f"{{{XM}}}sqref", ""): item.findtext(f"{{{XM}}}f", "")
        for item in group.iter(f"{{{X14}}}sparkline")
    }
    return SparklineInfo.model_construct(
        type=types.get(group.get("type", ""), "line"),
        colors=_read_colors(group),
        show=[point for point in _POINTS if flag(point)],
        show_axis=flag("displayXAxis"),
        axis_min=scale("min"),
        axis_max=scale("max"),
        right_to_left=flag("rightToLeft"),
        dates=None if dates is None or not flag("dateAxis") else dates.text,
        empty_cells=empty.get(group.get("displayEmptyCellsAs", ""), "zero"),
        hidden=flag("displayHidden"),
        line_weight=float(group.get("lineWeight", LINE_WEIGHT)),
        sparklines=sparklines,
    )


def _read_colors(group: ElementTree.Element) -> SparklineColors:
    """The colors given as RGB; theme colors keep Excel's default."""
    found = {}
    for name in _COLORS:
        element = group.find(f"{{{X14}}}color{name.capitalize()}")
        if element is not None and (rgb := element.get("rgb")):
            found[name] = rgb[2:]
    return SparklineColors(**found)
