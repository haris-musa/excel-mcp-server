"""Reading the cells of an openpyxl worksheet as calculator values."""

import datetime as dt
from collections.abc import Iterator
from decimal import Decimal
from typing import Any

from openpyxl.worksheet.worksheet import Worksheet

from excel_mcp.calc.values import ERRORS, VALUE, Scalar, UncalculableError, date_to_serial
from excel_mcp.lazy_workbook import LazyCells

CellKey = tuple[str, int, int]


def key(position: tuple[Worksheet, int, int]) -> CellKey:
    return (position[0].title, position[1], position[2])


def stored_in(
    sheet: Worksheet, top: int, left: int, bottom: int, right: int
) -> Iterator[tuple[tuple[int, int], Any]]:
    """The stored cells inside a rectangle, whichever is fewer to walk: it or the sheet."""
    if isinstance(sheet._cells, LazyCells):
        yield from sheet._cells.stored_in(top, left, bottom, right)
    elif (bottom - top + 1) * (right - left + 1) <= len(sheet._cells):
        for row in range(top, bottom + 1):
            for col in range(left, right + 1):
                if (cell := sheet._cells.get((row, col))) is not None:
                    yield (row, col), cell
    else:
        for (row, col), cell in sheet._cells.items():
            if top <= row <= bottom and left <= col <= right:
                yield (row, col), cell


def constant(cell: Any) -> Scalar:
    value = cell.value
    if cell.data_type == "e":
        return ERRORS.get(str(value), VALUE)
    match value:
        case dt.datetime():
            return (
                date_to_serial(value.date())
                + (value - dt.datetime.combine(value.date(), dt.time())).total_seconds() / 86400
            )
        case dt.date():
            return date_to_serial(value)
        case dt.time():
            return (value.hour * 3600 + value.minute * 60 + value.second) / 86400
        case dt.timedelta():
            return value.total_seconds() / 86400
        case Decimal():
            return float(value)
        case bool() | int() | float() | str() | None:
            return value
    raise UncalculableError("unsupported cell content")
