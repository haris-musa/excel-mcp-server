"""Finding and removing the PivotTables already in a workbook."""

from collections.abc import Iterator

from openpyxl.pivot.table import TableDefinition
from openpyxl.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet
from pydantic import BaseModel

from excel_mcp.errors import InvalidArgumentError
from excel_mcp.operations.cells import clear_range
from excel_mcp.package.guards import check_pivot_removal
from excel_mcp.refs import CellRange, parse_range
from excel_mcp.text import quoted
from excel_mcp.workspace import worksheets


class PivotInfo(BaseModel):
    name: str
    range: str
    source: str | None = None


def sheet_pivots(sheet: Worksheet) -> list[TableDefinition]:
    return sheet._pivots  # pyright: ignore[reportAttributeAccessIssue]


def workbook_pivots(workbook: Workbook) -> Iterator[TableDefinition]:
    for sheet in worksheets(workbook):
        yield from sheet_pivots(sheet)


def pivot_area(pivot: TableDefinition) -> CellRange:
    """The cells the PivotTable occupies, including its filter fields above the table."""
    body = parse_range(pivot.location.ref)
    filter_rows = pivot.location.rowPageCount
    top = body.min_row - filter_rows - 1 if filter_rows else body.min_row
    return CellRange(top, body.min_col, body.max_row, body.max_col)


def list_pivots(sheet: Worksheet) -> list[PivotInfo]:
    return [
        PivotInfo(name=pivot.name, range=pivot.location.ref, source=_source(pivot))
        for pivot in sheet_pivots(sheet)
    ]


def delete_pivot(sheet: Worksheet, name: str, max_cells: int) -> PivotInfo:
    """Remove a PivotTable and the cells it filled."""
    pivots = sheet_pivots(sheet)
    for pivot in pivots:
        if pivot.name.casefold() == name.strip().casefold():
            check_pivot_removal(sheet, pivot.name)
            info = PivotInfo(name=pivot.name, range=pivot.location.ref, source=_source(pivot))
            pivots.remove(pivot)
            clear_range(
                sheet, str(pivot_area(pivot)), contents=True, formats=False, max_cells=max_cells
            )
            return info
    if not pivots:
        raise InvalidArgumentError(f"Sheet {sheet.title!r} has no PivotTables.")
    raise InvalidArgumentError(
        f"Sheet {sheet.title!r} has no PivotTable {name!r}. "
        f"Available: {quoted(pivot.name for pivot in pivots)}."
    )


def _source(pivot: TableDefinition) -> str | None:
    source = pivot.cache.cacheSource.worksheetSource if pivot.cache else None
    if source is None:
        return None
    if source.sheet and source.ref:
        return f"{source.sheet}!{source.ref}"
    return source.name
