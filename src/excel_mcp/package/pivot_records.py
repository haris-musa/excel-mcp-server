"""Pointing each PivotTable cache at its own records when the file is saved."""

from openpyxl import Workbook
from openpyxl.pivot.cache import CacheDefinition
from openpyxl.worksheet.worksheet import Worksheet


def prepare(workbook: Workbook) -> None:
    """Give each cache's records the number of the cache.

    openpyxl links a cache to its records before it gives them their number, so without this
    every cache after the first is linked to the first cache's records.
    """
    caches: set[CacheDefinition] = set()
    for sheet in workbook.worksheets:
        if not isinstance(sheet, Worksheet):
            continue
        for pivot in sheet._pivots:  # pyright: ignore[reportAttributeAccessIssue]
            if pivot.cache not in caches:
                caches.add(pivot.cache)
                if pivot.cache.records is not None:
                    pivot.cache.records._id = len(caches)
