"""Putting what openpyxl dropped back into the file it wrote."""

import re

from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from excel_mcp.package import consistency, metadata, properties
from excel_mcp.package.model import (
    Link,
    PackageState,
    Part,
    Sheet,
    state_of,
    workbook_sheets,
)
from excel_mcp.package.opc import Package, Rel, free_id
from excel_mcp.package.patch import WORKBOOK_ORDER, append_to_root, insertions, splice
from excel_mcp.package.restore_sheet import restore_sheet
from excel_mcp.package.scan import attributes, rel_id, remap_relationship_ids, scan, tags

MAIN = "xl/workbook.xml"
_TAB_ID = re.compile(r'(<(?:[\w.-]+:)?pivotTable\b[^>]*?\btabId=")(\d+)(")')
_TABLE_ID = re.compile(r'(<(?:[\w.-]+:)?tableSlicerCache\b[^>]*?\btableId=")(\d+)(")')


def apply(workbook: Workbook, written: bytes) -> bytes:
    """``written``, the file openpyxl saved for ``workbook``, with what it dropped put back."""
    state = state_of(workbook)
    if state.is_empty():
        return written
    return Restorer(workbook, state, written).run()


class Restorer:
    """Merges the preserved content of a `PackageState` into a package that openpyxl wrote."""

    def __init__(self, workbook: Workbook, state: PackageState, written: bytes) -> None:
        self.workbook = workbook
        self.state = state
        self.package = Package(written)
        self.placed: dict[int, str] = {}
        numbers = {sheet: number for number, sheet in enumerate(workbook_sheets(workbook), 1)}
        self.sheet_ids = {old: numbers[s] for s, old in state.sheet_ids.items() if s in numbers}
        self.table_ids = {old: table.id for table, old in state.table_ids.items()}
        self.pivots = consistency.pivots_that_remain(workbook)
        self.orphans = consistency.orphaned_caches(workbook, state)
        self.gone = {name for sheet, name in state.sheet_names.items() if sheet not in numbers}
        self.dynamic_array = 0

    def run(self) -> bytes:
        links = list(self.state.workbook.links)
        links += self._metadata(links)
        parts = self.sheet_parts()
        for sheet, package in self.state.sheets.items():
            if sheet in parts and isinstance(sheet, Worksheet):
                restore_sheet(self, sheet, package, parts[sheet])
        self._workbook(links)
        self._company(self.state.workbook.company)
        return self.package.to_bytes()

    def sheet_parts(self) -> dict[Sheet, str]:
        document = scan(self.package.read(MAIN))
        targets = {rel.id: rel.target for rel in self.package.rels(MAIN)}
        return {
            self.workbook[values["name"]]: targets[rel_id(values, document)]
            for values in tags(document, "sheet")
        }

    # -- parts and relationships -------------------------------------------------------------

    def place(self, part: Part) -> str:
        """Add a part, and the parts it refers to, to the package; its name may change."""
        if id(part) in self.placed:
            return self.placed[id(part)]
        data = self._with_current_ids(part)
        same = self.package.exists(part.name) and self.package.read(part.name) == data
        name = part.name if same else self.package.unique_name(part.name)
        self.placed[id(part)] = name
        if not same:
            self.package.write(name, data)
            self.package.content_types.declare(name, part.content_type, part.by_default)
            self.package.rels(name).extend(self._rels(part.links))
        return name

    def _rels(self, links: list[Link]) -> list[Rel]:
        return [
            Rel(link.id, link.type, self.target(link), isinstance(link.target, str))
            for link in links
        ]

    def target(self, link: Link) -> str:
        return link.target if isinstance(link.target, str) else self.place(link.target)

    def attach(
        self, owner: str, links: list[Link], texts: list[str], scope: dict[str, str]
    ) -> list[str]:
        """Add relationships to a part that is being merged into, and renumber the XML to match."""
        rels = self.package.rels(owner)
        mapping: dict[str, str] = {}
        for link in links:
            rel = Rel(link.id, link.type, self.target(link), isinstance(link.target, str))
            existing = next((r for r in rels if (r.type, r.target) == (rel.type, rel.target)), None)
            if existing is None:
                rel.id = free_id(rels, link.id)
                rels.append(rel)
            mapping[link.id] = (existing or rel).id
        return [remap_relationship_ids(text, mapping, scope) for text in texts]

    def _with_current_ids(self, part: Part) -> bytes:
        """Slicer and timeline caches name sheets and tables by number, which openpyxl renumbers."""
        if part.content_type not in (consistency.SLICER_CACHE, consistency.TIMELINE_CACHE):
            return part.data
        text = part.data.decode("utf-8")
        text = _TAB_ID.sub(lambda m: m[1] + str(self.sheet_ids.get(int(m[2]), m[2])) + m[3], text)
        text = _TABLE_ID.sub(lambda m: m[1] + str(self.table_ids.get(int(m[2]), m[2])) + m[3], text)
        return consistency.without_missing_pivots(text, self.pivots).encode("utf-8")

    # -- the workbook ------------------------------------------------------------------------

    def _metadata(self, links: list[Link]) -> list[Link]:
        """Add dynamic array metadata if a cell asks for it; returns the link if the part is new."""
        if not any(mark.dynamic for sheet in self.state.sheets.values() for mark in sheet.marks):
            return []
        existing = next((link for link in links if link.type == metadata.RELATIONSHIP), None)
        part = existing.target if existing else Part(metadata.PART, metadata.CONTENT_TYPE, b"")
        if isinstance(part, str):
            return []
        document = metadata.Metadata(part.data.decode("utf-8") if part.data else None)
        self.dynamic_array = document.dynamic_array_index()
        part.data = document.xml.encode("utf-8")
        return [] if existing else [Link("rIdMetadata", metadata.RELATIONSHIP, part)]

    def _workbook(self, links: list[Link]) -> None:
        package = self.state.workbook
        names = [name for name, _ in package.elements]
        links, texts = consistency.without_caches(
            links,
            [xml for _, xml in package.elements] + list(package.extensions.values()),
            self.orphans,
        )
        extensions = consistency.empty_cache_lists(
            dict(zip(package.extensions, texts[len(names) :], strict=True))
        )
        remapped = self.attach(
            MAIN, links, texts[: len(names)] + list(extensions.values()), package.namespaces
        )
        new = list(zip(names, remapped[: len(names)], strict=True))
        if extensions:
            new.append(("extLst", f"<extLst>{''.join(remapped[len(names) :])}</extLst>"))
        self._cache_extensions(package.cache_extensions)
        data = self.package.read(MAIN)
        self.package.write(MAIN, splice(data, insertions(scan(data), new, WORKBOOK_ORDER)))
        self.attach("", package.root_links, [], {})

    def _cache_extensions(self, extensions: dict[str, str]) -> None:
        if not extensions:
            return
        document = scan(self.package.read(MAIN))
        targets = {rel.id: rel.target for rel in self.package.rels(MAIN)}
        for values in tags(document, "pivotCache"):
            extension = extensions.get(values["cacheId"])
            if extension:
                target = targets[rel_id(values, document)]
                self.package.write(target, append_to_root(self.package.read(target), extension))

    def _company(self, company: str) -> None:
        if company:
            app = self.package.read(properties.APP_PART)
            self.package.write(properties.APP_PART, properties.named(app, company))

    def pivot_name(self, part: str) -> str:
        data = self.package.read(part)
        document = scan(data)
        return attributes(data[document.root_tag[0] : document.root_tag[1]].decode())["name"]
