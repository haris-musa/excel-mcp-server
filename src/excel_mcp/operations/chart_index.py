"""Finding and removing the charts already on a sheet."""

from openpyxl.chart import BarChart
from openpyxl.chart.title import Title
from openpyxl.drawing.spreadsheet_drawing import OneCellAnchor, TwoCellAnchor
from openpyxl.worksheet.worksheet import Worksheet
from pydantic import BaseModel

from excel_mcp.errors import InvalidArgumentError
from excel_mcp.refs import cell_name


class ChartInfo(BaseModel):
    index: int
    type: str
    title: str | None
    anchor: str | None


def list_charts(sheet: Worksheet) -> list[ChartInfo]:
    return [
        ChartInfo(
            index=index,
            type=_type_name(chart),
            title=_title_text(chart.title),
            anchor=_anchor_cell(chart.anchor),
        )
        for index, chart in enumerate(sheet._charts, start=1)  # pyright: ignore[reportAttributeAccessIssue]
    ]


def delete_chart(sheet: Worksheet, index: int) -> ChartInfo:
    charts = list_charts(sheet)
    if not 1 <= index <= len(charts):
        valid = f"1 to {len(charts)}" if charts else "none: the sheet has no charts"
        raise InvalidArgumentError(
            f"Sheet {sheet.title!r} has no chart {index}. Valid chart indices: {valid}. "
            "describe_sheet lists them."
        )
    del sheet._charts[index - 1]  # pyright: ignore[reportAttributeAccessIssue]
    return charts[index - 1]


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
