"""Sheet layout: column widths, row heights, frozen panes, filters and tab color."""

from openpyxl.utils.cell import column_index_from_string, get_column_letter
from openpyxl.worksheet.worksheet import Worksheet
from pydantic import BaseModel, Field

from excel_mcp.errors import InvalidArgumentError
from excel_mcp.operations.cells import stored_cells
from excel_mcp.operations.formatting import parse_color
from excel_mcp.refs import MAX_ROW, parse_cell, parse_range

MIN_AUTOFIT_WIDTH = 8
MAX_AUTOFIT_WIDTH = 80


class ColumnWidth(BaseModel):
    """The width of one column."""

    column: str = Field(description="Column letter, e.g. 'B'.")
    width: float = Field(gt=0, le=255, description="Width in characters.")


class RowHeight(BaseModel):
    """The height of one row."""

    row: int = Field(ge=1, le=MAX_ROW, description="1-based row number.")
    height: float = Field(gt=0, le=409, description="Height in points.")


class SheetLayout(BaseModel):
    """Layout changes. Fields left as null are not changed."""

    column_widths: list[ColumnWidth] | None = Field(default=None, description="Widths to set.")
    row_heights: list[RowHeight] | None = Field(default=None, description="Heights to set.")
    autofit_columns: list[str] | None = Field(
        default=None, description="Column letters to size to their content, e.g. ['A', 'C']."
    )
    freeze_panes: str | None = Field(
        default=None, description="First unfrozen cell: 'A2' freezes row 1, 'A1' unfreezes."
    )
    auto_filter: str | None = Field(
        default=None, description="Range with filter buttons, e.g. 'A1:F100'."
    )
    tab_color: str | None = Field(default=None, description="Sheet tab hex color.")


def apply_layout(sheet: Worksheet, layout: SheetLayout) -> None:
    for column_width in layout.column_widths or []:
        sheet.column_dimensions[column_letter(column_width.column)].width = column_width.width
    for row_height in layout.row_heights or []:
        sheet.row_dimensions[row_height.row].height = row_height.height
    for column in layout.autofit_columns or []:
        letter = column_letter(column)
        sheet.column_dimensions[letter].width = estimate_width(sheet, letter)
    if layout.freeze_panes is not None:
        parse_cell(layout.freeze_panes)
        sheet.freeze_panes = layout.freeze_panes.upper()
    if layout.auto_filter is not None:
        area = parse_range(layout.auto_filter)
        for table in sheet.tables.values():
            if area.overlaps(parse_range(table.ref)):
                raise InvalidArgumentError(
                    f"{area} overlaps table {table.displayName!r}, which has its own filter."
                )
        sheet.auto_filter.ref = str(area)
    if layout.tab_color is not None:
        sheet.sheet_properties.tabColor = parse_color(layout.tab_color)


def column_letter(column: str) -> str:
    try:
        return get_column_letter(column_index_from_string(column.strip().upper()))
    except ValueError:
        raise InvalidArgumentError(
            f"Invalid column {column!r}. Use a letter such as 'B'."
        ) from None


def estimate_width(sheet: Worksheet, letter: str) -> float:
    """Approximate a fitting width from the longest text line in the column."""
    index = column_index_from_string(letter)
    lines = [
        line
        for cell in stored_cells(sheet)
        if cell.column == index
        for line in str(cell.value).splitlines()
    ]
    longest = max((len(line) for line in lines), default=0)
    return min(max(longest + 2, MIN_AUTOFIT_WIDTH), MAX_AUTOFIT_WIDTH)
