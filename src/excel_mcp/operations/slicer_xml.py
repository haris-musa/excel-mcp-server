"""The XML of slicers and timelines, as Excel writes it, and how to read it back."""

import datetime as dt
import re
from dataclasses import dataclass
from xml.sax.saxutils import quoteattr

from excel_mcp.package.opc import XML_DECLARATION
from excel_mcp.package.scan import attributes

X14 = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
X15 = "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"
MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

SLICER_PART = "application/vnd.ms-excel.slicer+xml"
TIMELINE_PART = "application/vnd.ms-excel.timeline+xml"
SLICER_CACHE_REL = "http://schemas.microsoft.com/office/2007/relationships/slicerCache"
SLICER_REL = "http://schemas.microsoft.com/office/2007/relationships/slicer"
TIMELINE_CACHE_REL = "http://schemas.microsoft.com/office/2011/relationships/timelineCache"
TIMELINE_REL = "http://schemas.microsoft.com/office/2011/relationships/timeline"
PIVOT_CACHE_ID_EXTENSION = "{725AE2AE-9491-48be-B2B4-4EB974FC3084}"

# Extension lists: the workbook's, then the sheet's.
PIVOT_CACHES = "{BBE1A952-AA13-448e-AADC-164F8A28A991}"
WORKBOOK_PR = "{79F54976-1DA5-4618-B147-4CDE4B953A38}"
TABLE_CACHES = "{46BE6895-7355-4a93-B00E-2C351335B9C9}"
TIMELINE_CACHES = "{D0CA8CA8-9F24-4464-BF8E-62219DCF47F9}"
PIVOT_SLICERS = "{A8765BA9-456A-4dab-B4F3-ACF838C121DE}"
TABLE_SLICERS = "{3A4CF648-6AED-40f4-86FF-DC5316D8AED3}"
TIMELINES = "{7E03D99C-DC04-49d9-9315-930204A7B6E9}"
TABLE_CACHE_EXTENSION = "{2F2917AC-EB37-4324-AD4E-5DD8C200BD13}"
HIDE_EMPTY_EXTENSION = "{470722E0-AACD-4C17-9CDC-17EF765DBC7E}"

DEFAULT_ROW_HEIGHT = 241300
SCROLL_BACK_DAYS = 227
LEVELS = ("years", "quarters", "months", "days")

_ROOT = f'xmlns:mc="{MC}" mc:Ignorable="x" xmlns:x="{MAIN}"'


@dataclass(frozen=True)
class CacheOptions:
    sort: str = "ascending"
    hide_empty: bool = False


def _attrs(**values: object) -> str:
    return "".join(f" {k}={quoteattr(str(v))}" for k, v in values.items() if v is not None)


def _date(moment: dt.date | dt.datetime) -> str:
    return f"{dt.datetime.combine(moment, dt.time()):%Y-%m-%dT%H:%M:%S}"


def _hide_empty(options: CacheOptions) -> str:
    if not options.hide_empty:
        return ""
    return (
        f'<x:ext uri="{HIDE_EMPTY_EXTENSION}" xmlns:x15="{X15}">'
        "<x15:slicerCacheHideItemsWithNoData/></x:ext>"
    )


def pivot_cache(
    name: str,
    field: str,
    pivots: list[tuple[int, str]],
    cache_id: int,
    items: list[tuple[int, bool]],
    empty: set[int],
    options: CacheOptions,
) -> str:
    """A slicer cache over a PivotTable field; ``items`` are (item, selected) in order and
    ``empty`` the items no record falls in."""
    listed = "".join(
        f"<i{_attrs(x=x, s='1' if on else None, nd='1' if x in empty else None)}/>"
        for x, on in items
    )
    sort = _attrs(sortOrder="descending" if options.sort == "descending" else None)
    extension = f"<extLst>{_hide_empty(options)}</extLst>" if options.hide_empty else ""
    return (
        f"{_cache_head(name, field)}<pivotTables>{_pivot_entries(pivots)}</pivotTables>"
        f'<data><tabular pivotCacheId="{cache_id}"{sort}>'
        f'<items count="{len(items)}">{listed}</items>'
        f"</tabular></data>{extension}</slicerCacheDefinition>"
    )


def table_cache(name: str, field: str, table_id: int, column: int, options: CacheOptions) -> str:
    sort = _attrs(sortOrder="descending" if options.sort == "descending" else None)
    return (
        f"{_cache_head(name, field)}"
        f'<extLst><x:ext uri="{TABLE_CACHE_EXTENSION}" xmlns:x15="{X15}">'
        f'<x15:tableSlicerCache tableId="{table_id}" column="{column}"{sort}/></x:ext>'
        f"{_hide_empty(options)}</extLst></slicerCacheDefinition>"
    )


def _cache_head(name: str, field: str) -> str:
    return (
        f'{XML_DECLARATION}<slicerCacheDefinition xmlns="{X14}" {_ROOT}'
        f"{_attrs(name=name, sourceName=field)}>"
    )


def _pivot_entries(pivots: list[tuple[int, str]]) -> str:
    return "".join(f"<pivotTable{_attrs(tabId=tab, name=name)}/>" for tab, name in pivots)


def timeline_cache(
    name: str,
    field: str,
    pivots: list[tuple[int, str]],
    cache_id: int,
    bounds: tuple[dt.date, dt.date],
    selection: tuple[dt.date, dt.date] | None,
) -> str:
    chosen = (
        f'<selection startDate="{_date(selection[0])}" endDate="{_date(selection[1])}"/>'
        if selection
        else ""
    )
    return (
        f'{XML_DECLARATION}<timelineCacheDefinition xmlns="{X15}" xmlns:x15="{X15}" '
        f'xmlns:mc="{MC}"{_attrs(name=name, sourceName=field)}>'
        f"<pivotTables>{_pivot_entries(pivots)}</pivotTables>"
        f'<state minimalRefreshVersion="6" lastRefreshVersion="6" pivotCacheId="{cache_id}" '
        f'filterType="{"dateBetween" if selection else "unknown"}">{chosen}'
        f'<bounds startDate="{_date(bounds[0])}" endDate="{_date(bounds[1])}"/></state>'
        "</timelineCacheDefinition>"
    )


def slicer_entry(
    name: str, cache: str, caption: str, columns: int, header: bool, style: str | None
) -> str:
    return (
        "<slicer"
        + _attrs(
            name=name,
            cache=cache,
            caption=caption,
            columnCount=columns if columns > 1 else None,
            showCaption=None if header else 0,
            style=style,
            rowHeight=DEFAULT_ROW_HEIGHT,
        )
        + "/>"
    )


def slicers_part(entries: str) -> str:
    return f'{XML_DECLARATION}<slicers xmlns="{X14}" {_ROOT}>{entries}</slicers>'


def timeline_entry(
    name: str, cache: str, caption: str, level: int, scroll: dt.date, style: str | None
) -> str:
    return (
        "<timeline"
        + _attrs(
            name=name,
            cache=cache,
            caption=caption,
            level=level,
            selectionLevel=2,
            scrollPosition=_date(scroll),
            style=style,
        )
        + "/>"
    )


def timelines_part(entries: str) -> str:
    return f'{XML_DECLARATION}<timelines xmlns="{X15}" {_ROOT}>{entries}</timelines>'


def timeline_bounds(first: dt.date, last: dt.date) -> tuple[dt.date, dt.date]:
    """From the start of the first date's year to the start of the year after the last."""
    return dt.date(first.year, 1, 1), dt.date(last.year + 1, 1, 1)


def timeline_scroll(bounds: tuple[dt.date, dt.date]) -> dt.date:
    return max(bounds[0], bounds[1] - dt.timedelta(days=SCROLL_BACK_DAYS))


# -- extension lists -------------------------------------------------------------------------


def workbook_cache_extensions(kind: str, rid: str) -> tuple[str, str]:
    """The workbook ``<ext>`` that lists a new cache: (uri, XML with just that cache)."""
    ref = f'<x14:slicerCache xmlns:r="{R}" r:id="{rid}"/>'
    if kind == "pivot":
        return (
            PIVOT_CACHES,
            f'<ext uri="{PIVOT_CACHES}" xmlns:x14="{X14}">'
            f"<x14:slicerCaches>{ref}</x14:slicerCaches></ext>",
        )
    if kind == "table":
        return TABLE_CACHES, (
            f'<ext uri="{TABLE_CACHES}" xmlns:x15="{X15}"><x15:slicerCaches xmlns:x14="{X14}">'
            f"{ref}</x15:slicerCaches></ext>"
        )
    return TIMELINE_CACHES, (
        f'<ext uri="{TIMELINE_CACHES}" xmlns:x15="{X15}"><x15:timelineCacheRefs>'
        f'<x15:timelineCacheRef xmlns:r="{R}" r:id="{rid}"/></x15:timelineCacheRefs></ext>'
    )


def sheet_slicer_extension(kind: str, rid: str) -> tuple[str, str]:
    """The sheet ``<ext>`` that lists a slicers or timelines part."""
    if kind == "pivot":
        return PIVOT_SLICERS, (
            f'<ext uri="{PIVOT_SLICERS}" xmlns:x14="{X14}"><x14:slicerList>'
            f'<x14:slicer xmlns:r="{R}" r:id="{rid}"/></x14:slicerList></ext>'
        )
    if kind == "table":
        return TABLE_SLICERS, (
            f'<ext uri="{TABLE_SLICERS}" xmlns:x15="{X15}"><x14:slicerList xmlns:x14="{X14}">'
            f'<x14:slicer xmlns:r="{R}" r:id="{rid}"/></x14:slicerList></ext>'
        )
    return TIMELINES, (
        f'<ext uri="{TIMELINES}" xmlns:x15="{X15}"><x15:timelineRefs>'
        f'<x15:timelineRef xmlns:r="{R}" r:id="{rid}"/></x15:timelineRefs></ext>'
    )


WORKBOOK_PR_EXTENSION = f'<ext uri="{WORKBOOK_PR}" xmlns:x14="{X14}"><x14:workbookPr/></ext>'


def cache_id_extension(cache_id: int) -> str:
    return (
        f'<extLst><ext uri="{PIVOT_CACHE_ID_EXTENSION}" xmlns:x14="{X14}">'
        f'<x14:pivotCacheDefinition pivotCacheId="{cache_id}"/></ext></extLst>'
    )


# -- reading ---------------------------------------------------------------------------------


@dataclass(frozen=True)
class CacheInfo:
    name: str
    field: str
    kind: str
    pivots: list[tuple[int, str]]
    table_id: int | None
    column: int | None
    items: list[tuple[int, bool]]
    sort: str
    hide_empty: bool
    selection: tuple[dt.datetime, dt.datetime] | None


def read_cache(xml: str) -> CacheInfo:
    values = attributes(re.search(r"<(?:slicer|timeline)CacheDefinition\b[^>]*>", xml)[0])  # pyright: ignore[reportOptionalSubscript]
    pivots = [
        (int(a["tabId"]), a["name"])
        for a in (attributes(t) for t in re.findall(r"<pivotTable\b[^>]*>", xml))
    ]
    items = [
        (int(a["x"]), a.get("s") == "1")
        for a in (attributes(t) for t in re.findall(r"<i\b[^>]*>", xml))
    ]
    table = re.search(r"<x15:tableSlicerCache\b[^>]*>", xml)
    tabular = re.search(r"<tabular\b[^>]*>", xml)
    timeline = "<timelineCacheDefinition" in xml
    chosen = re.search(r"<selection\b[^>]*>", xml)
    detail = attributes(table[0]) if table else {}
    sort = (attributes(tabular[0]) if tabular else detail).get("sortOrder", "ascending")
    return CacheInfo(
        name=values["name"],
        field=values["sourceName"],
        kind="timeline" if timeline else "table" if table else "pivot",
        pivots=pivots,
        table_id=int(detail["tableId"]) if table else None,
        column=int(detail["column"]) if table else None,
        items=items,
        sort=sort,
        hide_empty="slicerCacheHideItemsWithNoData" in xml,
        selection=_selection(chosen[0]) if chosen else None,
    )


def _selection(tag: str) -> tuple[dt.datetime, dt.datetime]:
    values = attributes(tag)
    return dt.datetime.fromisoformat(values["startDate"]), dt.datetime.fromisoformat(
        values["endDate"]
    )


def read_entries(xml: str) -> list[dict[str, str]]:
    """The attributes of each ``<slicer>`` or ``<timeline>`` of a part."""
    return [attributes(t) for t in re.findall(r"<(?:slicer|timeline)\b[^>]*>", xml)]


def relationship_type(kind: str) -> str:
    return TIMELINE_REL if kind == "timeline" else SLICER_REL
