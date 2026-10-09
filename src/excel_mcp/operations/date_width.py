"""Widening columns for dates, as Excel does when a date is entered in a default-width column."""

from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from excel_mcp.values import DATE_FORMAT, DATETIME_FORMAT

# Excel's fitted widths for the formats `values.date_number_format` applies, in the units of
# the file (the column width plus 5 pixels of padding, in 7-pixel characters).
FITTED_WIDTHS = {DATE_FORMAT: 10.42578125, DATETIME_FORMAT: 18.28515625}
_DEFAULT_WIDTH = 9.140625


def widen_for_dates(sheet: Worksheet, needed: dict[int, float]) -> None:
    """Widen the columns (by index) that have no width of their own to the one they need.

    A column that has a width is left alone, and no column is made narrower.
    """
    sized = set()
    for dimension in sheet.column_dimensions.values():
        first = dimension.min or column_index_from_string(dimension.index)
        sized.update(range(first, (dimension.max or first) + 1))
    default = sheet.sheet_format.defaultColWidth or _DEFAULT_WIDTH
    for index, width in needed.items():
        if index not in sized and width > default:
            sheet.column_dimensions[get_column_letter(index)].width = width
