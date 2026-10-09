"""Finding and removing the charts already on a sheet, and naming its charts and pictures."""

from openpyxl.chart import BarChart
from openpyxl.chart._chart import ChartBase
from openpyxl.chart.title import Title
from openpyxl.worksheet.worksheet import Worksheet

from excel_mcp.operations import chartex, drawings
from excel_mcp.operations.chart_info import ChartInfo
from excel_mcp.package.shape_names import name_of
from excel_mcp.refs import parse_cell


def name_shapes(sheet: Worksheet) -> None:
    """Give each chart and picture of the sheet its own name."""
    taken = {chart.name.casefold() for chart in chartex.modern_charts(sheet)}
    drawings.name_uniquely(sheet._charts, "Chart", taken)  # pyright: ignore[reportAttributeAccessIssue]
    drawings.name_uniquely(sheet._images, "Picture", taken)  # pyright: ignore[reportAttributeAccessIssue]


def shape_names(sheet: Worksheet) -> set[str]:
    """The casefolded names of the sheet's charts and pictures."""
    name_shapes(sheet)
    shapes = [*sheet._charts, *sheet._images]  # pyright: ignore[reportAttributeAccessIssue]
    names = [name_of(shape) or "" for shape in shapes]
    return {name.casefold() for name in [*names, *(c.name for c in chartex.modern_charts(sheet))]}


def list_charts(sheet: Worksheet) -> list[ChartInfo]:
    """The charts of a sheet: those openpyxl models first, then the Excel 2016 ones."""
    name_shapes(sheet)
    classic = [
        ChartInfo(
            name=name_of(chart) or "",
            type=_type_name(chart),
            title=_title_text(chart.title),
            range=chart_range(sheet, chart),
            series=[_values_ref(item) for plot in chart._charts for item in plot.series],
        )
        for chart in sheet._charts  # pyright: ignore[reportAttributeAccessIssue]
    ]
    return classic + chartex.describe(sheet)


def chart_range(sheet: Worksheet, chart: ChartBase) -> str | None:
    """The cells a chart covers; a chart not saved yet has only its first cell and size."""
    if isinstance(chart.anchor, str):
        row, col = parse_cell(chart.anchor)
        size = (round(chart.width * drawings.EMU_PER_CM), round(chart.height * drawings.EMU_PER_CM))
        return drawings.extent(sheet, row, col, *size)
    return drawings.covered(sheet, chart.anchor)


def find_chart(sheet: Worksheet, name: str) -> int:
    """The 1-based position of the chart with this name among the sheet's charts."""
    names = [chart.name for chart in list_charts(sheet)]
    return drawings.find(names, name, "chart", sheet) + 1


def delete_chart(sheet: Worksheet, name: str) -> ChartInfo:
    index = find_chart(sheet, name)
    removed = list_charts(sheet)[index - 1]
    classic = len(sheet._charts)  # pyright: ignore[reportAttributeAccessIssue]
    if index > classic:
        chartex.remove(sheet, chartex.modern_charts(sheet)[index - classic - 1])
    else:
        del sheet._charts[index - 1]  # pyright: ignore[reportAttributeAccessIssue]
    return removed


def replace_chart(sheet: Worksheet, index: int, chart: ChartBase, anchor: str | None) -> None:
    """Put `chart` where chart `index` was in the sheet's list, drawn at `anchor`."""
    if index > len(sheet._charts):  # pyright: ignore[reportAttributeAccessIssue]
        delete_chart(sheet, list_charts(sheet)[index - 1].name)
        sheet.add_chart(chart, anchor)
        return
    sheet.add_chart(chart, anchor)
    sheet._charts[index - 1] = sheet._charts.pop()  # pyright: ignore[reportAttributeAccessIssue]


def _values_ref(series) -> str:
    source = series.val or series.yVal
    return source.numRef.f if source and source.numRef else ""


def _type_name(chart) -> str:
    if isinstance(chart, BarChart):
        return "column" if chart.type == "col" else "bar"
    return type(chart).__name__.removesuffix("Chart").lower()


def _title_text(title: Title | None) -> str | None:
    if title is None or title.tx is None or title.tx.rich is None:
        return None
    runs = [run.t for paragraph in title.tx.rich.p for run in paragraph.r]
    return "".join(runs) or None
