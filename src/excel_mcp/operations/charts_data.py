"""Where a chart's numbers come from: series ranges, names and categories, on any sheet."""

from dataclasses import dataclass
from typing import Literal

from openpyxl import Workbook
from openpyxl.chart.data_source import StrRef
from openpyxl.chart.series import SeriesLabel
from openpyxl.utils import absolute_coordinate, quote_sheetname

from excel_mcp.errors import InvalidArgumentError
from excel_mcp.operations.charts_options import ChartType, SeriesSpec
from excel_mcp.refs import CellRange, parse_cell, parse_range
from excel_mcp.workspace import get_sheet

SeriesIn = Literal["columns", "rows"]


@dataclass(frozen=True)
class Plot:
    """One series with its references resolved to absolute, sheet-qualified formulas."""

    spec: SeriesSpec
    values: str
    categories: str | None
    sizes: str | None
    name: SeriesLabel | None
    points: int


def resolve_series(
    workbook: Workbook,
    sheet: str,
    chart_type: ChartType,
    data_range: str | None,
    series_in: SeriesIn,
    series: list[SeriesSpec],
    categories: str | None,
) -> list[Plot]:
    if (data_range is None) == (not series):
        raise InvalidArgumentError(
            "Pass either data_range (a block of data) or series (explicit ranges), not both "
            "and not neither."
        )
    if data_range is not None:
        series = _block_series(workbook, sheet, data_range, chart_type, series_in)
        categories = None
    return [_resolve(workbook, sheet, item, categories) for item in series]


def _resolve(workbook: Workbook, sheet: str, spec: SeriesSpec, categories: str | None) -> Plot:
    values, points = _line(workbook, sheet, spec.values, "values")
    labels = spec.categories or categories
    return Plot(
        spec=spec,
        values=values,
        categories=_line(workbook, sheet, labels, "categories")[0] if labels else None,
        sizes=_line(workbook, sheet, spec.sizes, "sizes")[0] if spec.sizes else None,
        name=_name(workbook, spec.name),
        points=points,
    )


def _split_sheet(workbook: Workbook, default: str, text: str) -> tuple[str, str]:
    """Split 'Data!B2:B9' (or 'My Sheet'!B2:B9) into the worksheet's title and the cells."""
    qualifier, mark, cells = text.rpartition("!")
    if not mark:
        return default, text
    name = qualifier.removeprefix("'").removesuffix("'").replace("''", "'")
    return get_sheet(workbook, name).title, cells


def _qualified(sheet: str, area: CellRange) -> str:
    return f"{quote_sheetname(sheet)}!{absolute_coordinate(str(area))}"


def _line(workbook: Workbook, default: str, text: str, what: str) -> tuple[str, int]:
    """A one-row or one-column range as an absolute formula, and its number of cells."""
    sheet, cells = _split_sheet(workbook, default, text)
    area = parse_range(cells)
    if area.rows > 1 and area.cols > 1:
        raise InvalidArgumentError(
            f"Series {what} must be one row or one column, got {text!r}. List one series per "
            "column, or use data_range for a block."
        )
    return _qualified(sheet, area), area.size


def _name(workbook: Workbook, text: str | None) -> SeriesLabel | None:
    """A reference when `text` is a sheet-qualified cell such as 'Data!B1', else a literal."""
    if text is None:
        return None
    qualifier, mark, cell = text.rpartition("!")
    sheet = qualifier.removeprefix("'").removesuffix("'").replace("''", "'")
    if mark and sheet in workbook.sheetnames:
        row, col = parse_cell(cell)
        return SeriesLabel(strRef=StrRef(f=_qualified(sheet, CellRange(row, col, row, col))))
    return SeriesLabel(v=text)


def _block_series(
    workbook: Workbook, default: str, data_range: str, chart_type: ChartType, series_in: SeriesIn
) -> list[SeriesSpec]:
    """Series for a block with a header line and a label line, by columns or by rows."""
    sheet, cells = _split_sheet(workbook, default, data_range)
    area = parse_range(cells)
    if area.rows < 2 or area.cols < 2:
        raise InvalidArgumentError(
            "data_range needs a header row and a label column plus at least one series, "
            "e.g. 'A1:C10' with labels in A and series in B and C."
        )
    by_columns = series_in == "columns"
    lines = area.cols if by_columns else area.rows
    points = (area.rows if by_columns else area.cols) - 1

    def segment(line: int, start: int, stop: int) -> str:
        """Cells `start` to `stop` along line number `line`, the label line being 0."""
        if by_columns:
            cells = CellRange(
                area.min_row + start, area.min_col + line, area.min_row + stop, area.min_col + line
            )
        else:
            cells = CellRange(
                area.min_row + line, area.min_col + start, area.min_row + line, area.min_col + stop
            )
        return _qualified(sheet, cells)

    if chart_type == "bubble" and lines != 3:
        raise InvalidArgumentError(
            f"A bubble data_range has three {series_in}: x values, y values and sizes."
        )
    if chart_type == "pie" and lines > 2:
        raise InvalidArgumentError(f"A pie chart plots one series; use two {series_in} of data.")
    sizes = chart_type == "bubble"
    return [
        SeriesSpec(
            values=segment(line, 1, points),
            name=segment(line, 0, 0),
            categories=segment(0, 1, points),
            sizes=segment(2, 1, points) if sizes else None,
        )
        for line in ([1] if sizes else range(1, lines))
    ]
