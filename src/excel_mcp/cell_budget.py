"""A limit on the cells that a workbook's sheets make openpyxl create.

Memory use is set by the number of cells created, not by the size of the file: a sheet of
millions of empty cells is a few hundred kilobytes. So every sheet is counted before it is
parsed, and what is created cannot pass the limit. Rows count as cells (openpyxl keeps their
attributes), and so does every cell of a merged range or hyperlink range, which it creates
from the one element that names the range.
"""

from pathlib import Path
from typing import IO
from xml.parsers import expat

from openpyxl.reader.excel import ExcelReader
from openpyxl.xml.constants import SHEET_MAIN_NS, WORKSHEET_TYPE

from excel_mcp.errors import LimitExceededError
from excel_mcp.refs import parse_range

_CHUNK = 1 << 20
_NAMESPACE = f"{SHEET_MAIN_NS} "
_COUNTED = {f"{_NAMESPACE}c", f"{_NAMESPACE}row"}
_MERGED = f"{_NAMESPACE}mergeCell"
_LINKED = f"{_NAMESPACE}hyperlink"
# A cell that holds a link takes about three times the memory of one that does not.
_LINK_COST = 3


class CellBudget:
    """The cells that may still be created for one workbook, shared by all its sheets."""

    def __init__(self, limit: int) -> None:
        self.limit = limit
        self.used = 0

    def spend(self, count: int) -> None:
        self.used += count
        if self.used > self.limit:
            raise LimitExceededError(
                f"The workbook has more than {self.limit:,} cells and rows to load, which "
                "would use too much memory. Read it with read_range, split it into smaller "
                "workbooks, or raise --max-loaded-cells."
            )


def spend_sheet(source: IO[bytes], budget: CellBudget) -> None:
    """Spend the cells and rows of one worksheet's XML, stopping at the first one over the limit."""
    parser = expat.ParserCreate(namespace_separator=" ")

    def start(name: str, attributes: dict[str, str]) -> None:
        if name in _COUNTED:
            budget.spend(1)
        elif name == _MERGED:
            budget.spend(parse_range(attributes["ref"]).size)
        elif name == _LINKED:
            budget.spend(_LINK_COST * parse_range(attributes["ref"]).size)

    parser.StartElementHandler = start
    while chunk := source.read(_CHUNK):
        parser.Parse(chunk, False)
    parser.Parse(b"", True)


def check_workbook(path: Path, budget: CellBudget) -> None:
    """Spend the cells and rows that loading all of a workbook's worksheets would create."""
    reader = ExcelReader(path)
    try:
        reader.read_manifest()
        for part in reader.package.findall(WORKSHEET_TYPE):
            with reader.archive.open(part.PartName.lstrip("/")) as source:
                spend_sheet(source, budget)
    finally:
        reader.archive.close()
