"""Keeping preserved content consistent with what the edit did to the workbook.

Excel removes what only exists for something that is gone: slicers of a deleted sheet go
with their caches, a threaded comment whose note was deleted goes. A file that kept them
would be one Excel has to repair.
"""

import re
from dataclasses import replace

from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from excel_mcp.package.model import Link, PackageState, Part, Sheet, state_of, workbook_sheets
from excel_mcp.package.patch import Edit, splice
from excel_mcp.package.scan import attributes, scan

SLICER_CACHE = "application/vnd.ms-excel.slicerCache+xml"
TIMELINE_CACHE = "application/vnd.ms-excel.timelineCache+xml"
THREADED_COMMENTS = "application/vnd.ms-excel.threadedcomments+xml"
_CACHES = {SLICER_CACHE, TIMELINE_CACHE}
_SLICERS = {"application/vnd.ms-excel.slicer+xml", "application/vnd.ms-excel.timeline+xml"}
_CACHE_LISTS = {
    "{BBE1A952-AA13-448e-AADC-164F8A28A991}",  # slicer caches
    "{46BE6895-7355-4a93-B00E-2C351335B9C9}",  # slicer caches of Excel 2013
    "{D0CA8CA8-9F24-4464-BF8E-62219DCF47F9}",  # timeline caches
}
_SLICER_REFERENCE = re.compile(r'<(?:[\w.-]+:)?(?:slicer|timeline)\b[^>]*?\bcache="([^"]*)"')
PIVOT_ENTRY = re.compile(r"<(?:[\w.-]+:)?pivotTable\b[^>]*/>")
_CACHE_ENTRY = r'<(?:[\w.-]+:)?(?:slicerCache|timelineCacheRef)\b[^>]*?\b[\w.-]+:id="{id}"[^>]*/>'


def cache_name(part: Part) -> str:
    document = scan(part.data)
    return attributes(part.data[document.root_tag[0] : document.root_tag[1]].decode())["name"]


def slicer_caches(state: PackageState) -> dict[str, Link]:
    return {
        cache_name(link.target): link
        for link in state.workbook.links
        if isinstance(link.target, Part) and link.target.content_type in _CACHES
    }


def used_caches(workbook: Workbook, state: PackageState, leaving: Sheet | None = None) -> set[str]:
    """Names of the slicer and timeline caches that slicers on remaining sheets use.

    ``leaving`` is a sheet that is about to be deleted.
    """
    used: set[str] = set()
    for sheet, package in state.sheets.items():
        if sheet in workbook_sheets(workbook) and sheet is not leaving:
            for link in package.links:
                if isinstance(link.target, Part) and link.target.content_type in _SLICERS:
                    used |= set(_SLICER_REFERENCE.findall(link.target.data.decode("utf-8")))
    return used


def orphaned_caches(workbook: Workbook, state: PackageState) -> set[str]:
    """Slicer and timeline caches that no slicer uses any more; Excel deletes those."""
    return set(slicer_caches(state)) - used_caches(workbook, state)


def drop_orphaned_caches(workbook: Workbook) -> None:
    """Before saving: remove the defined names of caches that are going away."""
    for name in orphaned_caches(workbook, state_of(workbook)):
        if name in workbook.defined_names:
            del workbook.defined_names[name]


def without_caches(
    links: list[Link], texts: list[str], dropped: set[str]
) -> tuple[list[Link], list[str]]:
    """The workbook's links and extension entries, without the named slicer caches."""
    gone = {
        link.id
        for link in links
        if isinstance(link.target, Part)
        and link.target.content_type in _CACHES
        and cache_name(link.target) in dropped
    }
    kept = []
    for text in texts:
        for identifier in gone:
            text = re.sub(_CACHE_ENTRY.format(id=re.escape(identifier)), "", text)
        kept.append(text)
    return [link for link in links if link.id not in gone], kept


def empty_cache_lists(extensions: dict[str, str]) -> dict[str, str]:
    """Drop the lists of slicer caches that have no cache left."""
    return {
        uri: xml
        for uri, xml in extensions.items()
        if uri not in _CACHE_LISTS or re.search(r"[\w.-]+:id=\"", xml)
    }


def pivots_that_remain(workbook: Workbook) -> set[tuple[int, str]]:
    """(sheet number, name) of every PivotTable in the workbook as it is saved."""
    found = set()
    for number, sheet in enumerate(workbook_sheets(workbook), 1):
        if isinstance(sheet, Worksheet):
            found |= {(number, pivot.name) for pivot in sheet._pivots}  # pyright: ignore[reportAttributeAccessIssue]
    return found


def without_missing_pivots(xml: str, remaining: set[tuple[int, str]]) -> str:
    """Drop the connections of a slicer cache to PivotTables that were deleted."""

    def keep(match: re.Match[str]) -> str:
        values = attributes(match.group(0))
        return match.group(0) if (int(values["tabId"]), values["name"]) in remaining else ""

    return PIVOT_ENTRY.sub(keep, xml)


def reconcile_threads(part: Part, live: dict[str, str]) -> Part | None:
    """The threaded comments of a sheet, pointing at the notes they belong to.

    ``live`` maps the id of a thread to the cell of its note. A thread whose note was
    deleted goes; None means that none is left.
    """
    document = scan(part.data)
    edits: list[Edit] = []
    kept = 0
    for child in document.children:
        text = part.data[child.start : child.end]
        values = attributes(text[: text.index(b">") + 1].decode())
        thread = values.get("parentId") or values.get("id", "")
        if thread not in live:
            edits.append((child.start, child.end, b""))
            continue
        kept += 1
        moved = re.sub(rb'\bref="[^"]*"', f'ref="{live[thread]}"'.encode(), text, count=1)
        edits.append((child.start, child.end, moved))
    return replace(part, data=splice(part.data, edits)) if kept else None
