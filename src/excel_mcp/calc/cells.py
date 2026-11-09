"""Reading the cells of an openpyxl worksheet as calculator values."""

import datetime as dt
import math
from collections.abc import Iterator
from decimal import Decimal
from typing import Any

from openpyxl.worksheet.worksheet import Worksheet

from excel_mcp.calc.values import (
    ERRORS,
    NULL,
    VALUE,
    Scalar,
    UncalculableError,
    date_to_serial,
    flush,
)
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


def _double(number: int) -> float:
    try:
        return float(number)
    except OverflowError:
        raise UncalculableError("whole number beyond the largest double") from None


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
        case int() if not isinstance(value, bool):
            return _double(value)
        case float():
            return flush(value) if math.isfinite(value) else NULL  # Excel reads 1E+309 as #NULL!
        case bool() | str() | None:
            return value
    raise UncalculableError("unsupported cell content")
