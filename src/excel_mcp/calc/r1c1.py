"""R1C1 reference text (INDIRECT with a1=FALSE) to A1 reference text."""

import re

from openpyxl.utils import get_column_letter

from excel_mcp.calc.values import REF, FormulaError
from excel_mcp.refs import MAX_COLUMN, MAX_ROW

# A coordinate is absolute ("5") or relative to the formula cell ("[-2]"); none means "same".
_NUMBER = r"(?:\[(-?\d+)\]|(\d+))?"
_CELL = re.compile(rf"R{_NUMBER}C{_NUMBER}", re.IGNORECASE)
_ROW = re.compile(rf"R{_NUMBER}", re.IGNORECASE)
_COLUMN = re.compile(rf"C{_NUMBER}", re.IGNORECASE)


def to_a1(text: str, row: int, col: int) -> str:
    """Convert "R2C3", "R[1]C[-1]:R5C6", "R2:R4" or "C3" as seen from the cell at (row, col).

    Text that is not an R1C1 reference (a defined name) is returned unchanged.
    """
    parts = text.split(":")
    if len(parts) > 2:
        return text
    converted = [_part(part, row, col) for part in parts]
    if any(kind is None for kind, _ in converted):
        return text
    kinds = {kind for kind, _ in converted}
    if len(kinds) > 1:
        raise FormulaError(REF)
    texts = [part for _, part in converted]
    if len(texts) == 1 and kinds != {"cell"}:
        texts *= 2
    return ":".join(texts)


def _part(part: str, row: int, col: int) -> tuple[str | None, str]:
    if match := _CELL.fullmatch(part):
        return "cell", f"{_column(match, 3, col)}{_row(match, 1, row)}"
    if match := _ROW.fullmatch(part):
        return "row", str(_row(match, 1, row))
    if match := _COLUMN.fullmatch(part):
        return "column", _column(match, 1, col)
    return None, part


def _row(match: re.Match[str], group: int, here: int) -> int:
    return _coordinate(match, group, here, MAX_ROW)


def _column(match: re.Match[str], group: int, here: int) -> str:
    return get_column_letter(_coordinate(match, group, here, MAX_COLUMN))


def _coordinate(match: re.Match[str], group: int, here: int, limit: int) -> int:
    relative, absolute = match.group(group), match.group(group + 1)
    value = here + int(relative) if relative else int(absolute) if absolute else here
    if not 1 <= value <= limit:
        raise FormulaError(REF)
    return value
