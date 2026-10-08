"""Chart-wide formatting: size, title, legend, frame and plot area."""

# pyright: reportArgumentType=false, reportAttributeAccessIssue=false, reportOptionalMemberAccess=false
# openpyxl's stubs type its enum-like fields as Literals and its descriptors as plain attributes.

from openpyxl.chart._chart import ChartBase
from openpyxl.chart.legend import Legend
from openpyxl.chart.shapes import GraphicalProperties

from excel_mcp.operations import charts_look as look
from excel_mcp.operations.charts_options import ChartOptions, LegendPosition
from excel_mcp.operations.formatting import parse_color

_LEGEND_POSITIONS = {"right": "r", "left": "l", "top": "t", "bottom": "b"}
_DEFAULT_TITLE_SIZE = 14


def style_chart(chart: ChartBase, options: ChartOptions, legend: LegendPosition) -> None:
    chart.width = options.width_cm
    chart.height = options.height_cm
    chart.style = options.style
    # Without an explicit flag, Excel draws the chart with rounded corners.
    chart.roundedCorners = False
    chart.graphical_properties = GraphicalProperties(
        solidFill=look.scheme_color("bg1"), ln=look.thin_line(15000, 85000)
    )
    if options.title:
        size = (options.title_size or _DEFAULT_TITLE_SIZE) * 100
        chart.title = look.title(options.title, size)
    if legend == "none":
        chart.legend = None
    else:
        # Without an explicit overlay flag, Excel draws the legend over the plot.
        chart.legend = Legend(legendPos=_LEGEND_POSITIONS[legend], overlay=False)
        chart.legend.spPr = look.no_fill()
        chart.legend.txPr = look.text_properties(900)
    if options.plot_color:
        chart.plot_area.spPr = GraphicalProperties(solidFill=parse_color(options.plot_color)[2:])
    else:
        chart.plot_area.spPr = look.no_fill()
