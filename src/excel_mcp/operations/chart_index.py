"""Finding and removing the charts already on a sheet."""

from openpyxl.chart import BarChart
from openpyxl.chart._chart import ChartBase
from openpyxl.chart.title import Title
from openpyxl.drawing.spreadsheet_drawing import OneCellAnchor, TwoCellAnchor
from openpyxl.worksheet.worksheet import Worksheet

from excel_mcp.errors import InvalidArgumentError
from excel_mcp.operations import chartex
from excel_mcp.operations.chart_info import ChartInfo
from excel_mcp.refs import cell_name


def list_charts(sheet: Worksheet) -> list[ChartInfo]:
    """The charts of a sheet: those openpyxl models first, then the Excel 2016 ones."""
    classic = [
        ChartInfo(
            index=index,
            type=_type_name(chart),
            title=_title_text(chart.title),
            anchor=_anchor_cell(chart.anchor),
            series=[_values_ref(item) for plot in chart._charts for item in plot.series],
        )
        for index, chart in enumerate(sheet._charts, start=1)  # pyright: ignore[reportAttributeAccessIssue]
    ]
    return classic + chartex.describe(sheet, len(classic) + 1)


def delete_chart(sheet: Worksheet, index: int) -> ChartInfo:
    removed = check_index(sheet, index)
    classic = len(sheet._charts)  # pyright: ignore[reportAttributeAccessIssue]
    if index > classic:
        chartex.remove(sheet, chartex.modern_charts(sheet)[index - classic - 1])
    else:
        del sheet._charts[index - 1]  # pyright: ignore[reportAttributeAccessIssue]
    return removed


def replace_chart(sheet: Worksheet, index: int, chart: ChartBase, anchor: str | None) -> None:
    """Put `chart` where chart `index` was in the sheet's list, drawn at `anchor`."""
    check_index(sheet, index)
    if index > len(sheet._charts):  # pyright: ignore[reportAttributeAccessIssue]
        delete_chart(sheet, index)
        sheet.add_chart(chart, anchor)
        return
    sheet.add_chart(chart, anchor)
    sheet._charts[index - 1] = sheet._charts.pop()  # pyright: ignore[reportAttributeAccessIssue]


def check_index(sheet: Worksheet, index: int) -> ChartInfo:
    charts = list_charts(sheet)
    if not charts:
        raise InvalidArgumentError(f"Sheet {sheet.title!r} has no charts.")
    if not 1 <= index <= len(charts):
        raise InvalidArgumentError(
            f"Sheet {sheet.title!r} has no chart {index}. Valid chart indices: 1 to "
            f"{len(charts)}. describe_sheet lists them."
        )
    return charts[index - 1]


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


def _anchor_cell(anchor) -> str | None:
    if isinstance(anchor, OneCellAnchor | TwoCellAnchor):
        return cell_name(anchor._from.row + 1, anchor._from.col + 1)
    return None
