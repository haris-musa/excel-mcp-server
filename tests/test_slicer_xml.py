"""The XML this server writes for slicers is what Excel wrote in the fixture, apart from ids."""

import datetime as dt
import re
import zipfile

from excel_mcp.operations import slicer_xml as xml
from tests.package_support import FIXTURES

_NOISE = [
    r' xmlns:xr10="[^"]*"',
    r' xr10:uid="\{[^}]*\}"',
    r' mc:Ignorable="[^"]*"',
    r"<\?xml[^>]*\?>\s*",
    r' xmlns:mc="[^"]*"',
    r' rowHeight="\d+"',
    r' xmlns:x="[^"]*"',
]


def _excel(name: str) -> str:
    with zipfile.ZipFile(FIXTURES / "excel_slicers.xlsx") as archive:
        data = archive.read(name).decode("utf-8")
    return _plain(data)


def _plain(data: str) -> str:
    for noise in _NOISE:
        data = re.sub(noise, "", data)
    return data.strip()


def test_a_pivot_slicer_cache_is_what_excel_wrote() -> None:
    ours = xml.pivot_cache(
        "Slicer_Product",
        "Product",
        [(2, "PivotSales")],
        122376683,
        [(1, True), (2, True), (0, True)],
        xml.CacheOptions(),
    )
    assert _plain(ours) == _excel("xl/slicerCaches/slicerCache1.xml")


def test_a_table_slicer_cache_is_what_excel_wrote() -> None:
    ours = xml.table_cache("Slicer_Region", "Region", 1, 1, xml.CacheOptions())
    assert _plain(ours) == _excel("xl/slicerCaches/slicerCache3.xml")


def test_a_timeline_cache_is_what_excel_wrote() -> None:
    ours = xml.timeline_cache(
        "NativeTimeline_Date",
        "Date",
        [(2, "PivotSales")],
        122376683,
        (dt.date(2025, 1, 1), dt.date(2026, 1, 1)),
        None,
    )
    assert _plain(ours) == _excel("xl/timelineCaches/timelineCache1.xml")


def test_slicer_and_timeline_entries_are_what_excel_wrote() -> None:
    entry = xml.slicer_entry("Product", "Slicer_Product", "Product", 1, True, None)
    assert _plain(xml.slicers_part(entry)) == _excel("xl/slicers/slicer1.xml")
    shown = xml.timeline_entry("Date", "NativeTimeline_Date", "Date", 2, dt.date(2025, 5, 19), None)
    assert _plain(xml.timelines_part(shown)) == _excel("xl/timelines/timeline1.xml")


def test_the_bounds_and_scroll_position_follow_the_dates() -> None:
    bounds = xml.timeline_bounds(dt.date(2025, 1, 10), dt.date(2025, 12, 25))
    assert bounds == (dt.date(2025, 1, 1), dt.date(2026, 1, 1))
    assert xml.timeline_scroll(bounds) == dt.date(2025, 5, 19)
    assert xml.timeline_bounds(dt.date(2015, 6, 10), dt.date(2025, 8, 20)) == (
        dt.date(2015, 1, 1),
        dt.date(2026, 1, 1),
    )
