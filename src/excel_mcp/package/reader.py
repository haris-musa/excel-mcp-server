"""The original file of a workbook, read as parts with the relationships between them."""

import zipfile
from collections.abc import Callable

from excel_mcp.errors import LimitExceededError
from excel_mcp.package.model import Link, Part
from excel_mcp.package.opc import CONTENT_TYPES, ContentTypes, Rel, parse_rels, rels_name


class Reader:
    """The original file: its parts as `Part` objects, read at most once each."""

    def __init__(self, archive: zipfile.ZipFile, budget: int) -> None:
        self.archive = archive
        self.names = set(archive.namelist())
        self.types = ContentTypes(archive.read(CONTENT_TYPES))
        self.parts: dict[str, Part] = {}
        self.budget = budget

    def read(self, name: str) -> bytes:
        size = self.archive.getinfo(name).file_size
        if size > self.budget:
            raise LimitExceededError(
                "The workbook holds more content that is not spreadsheet data than the "
                f"{self.budget:,} bytes this server can carry along."
            )
        self.budget -= size
        return self.archive.read(name)

    def rels(self, source: str) -> list[Rel]:
        name = rels_name(source)
        return parse_rels(source, self.archive.read(name)) if name in self.names else []

    def part(self, name: str) -> Part | None:
        """A part with everything it refers to, or None if the file lacks it."""
        if name not in self.names:
            return None
        if name not in self.parts:
            content_type = self.types.of(name) or "application/octet-stream"
            by_default = name not in self.types.overrides
            part = self.parts[name] = Part(
                name, content_type, self.read(name), by_default=by_default
            )
            part.links = self.links(name, lambda _: True)
        return self.parts[name]

    def links(self, source: str, keep: Callable[[Rel], bool]) -> list[Link]:
        links = []
        for rel in self.rels(source):
            if not keep(rel):
                continue
            if rel.external:
                links.append(Link(rel.id, rel.type, rel.target))
            elif (part := self.part(rel.target)) is not None:
                links.append(Link(rel.id, rel.type, part))
        return links
