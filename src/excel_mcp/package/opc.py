"""The Open Packaging Conventions: relationships, content types and the zip container."""

import io
import posixpath
import re
import zipfile
from dataclasses import dataclass
from xml.sax.saxutils import quoteattr

from excel_mcp.errors import WorkbookError
from excel_mcp.package.scan import attributes, scan

RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
REL_BASE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
CONTENT_TYPES = "[Content_Types].xml"
XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'


@dataclass
class Rel:
    """A relationship; ``target`` is a part name, or a URL when ``external``."""

    id: str
    type: str
    target: str
    external: bool = False


def rels_name(part: str) -> str:
    """The part holding the relationships of ``part`` ('' is the package itself)."""
    directory, _, name = part.rpartition("/")
    return f"{directory}/_rels/{name}.rels" if directory else f"_rels/{name}.rels"


def resolve(source: str, target: str) -> str:
    """The part name a relationship target points to, relative to its source part."""
    if target.startswith("/"):
        name = posixpath.normpath(target[1:])
    else:
        name = posixpath.normpath(posixpath.join(posixpath.dirname(source), target))
    if name.startswith("../") or name == "..":
        raise WorkbookError(f"A relationship of {source or 'the package'} leaves the package.")
    return name


def parse_rels(source: str, data: bytes) -> list[Rel]:
    document = scan(data)
    rels = []
    for child in document.children:
        values = attributes(document.raw(child))
        external = values.get("TargetMode", "").casefold() == "external"
        target = values["Target"]
        rels.append(
            Rel(
                values["Id"],
                values["Type"],
                target if external else resolve(source, target),
                external,
            )
        )
    return rels


def write_rels(rels: list[Rel]) -> bytes:
    lines = [f'<Relationships xmlns="{RELS_NS}">']
    for rel in rels:
        target = rel.target if rel.external else "/" + rel.target
        mode = ' TargetMode="External"' if rel.external else ""
        lines.append(
            f"<Relationship Id={quoteattr(rel.id)} Type={quoteattr(rel.type)} "
            f"Target={quoteattr(target)}{mode}/>"
        )
    lines.append("</Relationships>")
    return (XML_DECLARATION + "".join(lines)).encode()


def free_id(rels: list[Rel], wanted: str) -> str:
    """``wanted`` if it is a free ``rIdN``, else the first free one."""
    taken = {rel.id for rel in rels}
    if re.fullmatch(r"rId\d+", wanted) and wanted not in taken:
        return wanted
    number = 1
    while f"rId{number}" in taken:
        number += 1
    return f"rId{number}"


class ContentTypes:
    """The ``[Content_Types].xml`` of a package."""

    def __init__(self, data: bytes) -> None:
        self.defaults: dict[str, str] = {}
        self.overrides: dict[str, str] = {}
        document = scan(data)
        for child in document.children:
            values = attributes(document.raw(child))
            if child.local == "Default":
                self.defaults[values["Extension"].casefold()] = values["ContentType"]
            elif child.local == "Override":
                self.overrides[values["PartName"].lstrip("/")] = values["ContentType"]

    def of(self, part: str) -> str | None:
        extension = posixpath.splitext(part)[1].lstrip(".").casefold()
        return self.overrides.get(part) or self.defaults.get(extension)

    def declare(self, part: str, content_type: str, by_default: bool = False) -> None:
        """Make ``part`` have ``content_type``; ``by_default`` files its extension under it."""
        extension = posixpath.splitext(part)[1].lstrip(".").casefold()
        if by_default and extension and extension not in self.defaults:
            self.defaults[extension] = content_type
        if self.of(part) != content_type:
            self.overrides[part] = content_type

    def to_bytes(self) -> bytes:
        lines = [f'<Types xmlns="{TYPES_NS}">']
        for extension, content_type in self.defaults.items():
            lines.append(
                f"<Default Extension={quoteattr(extension)} ContentType={quoteattr(content_type)}/>"
            )
        for part, content_type in self.overrides.items():
            lines.append(
                f"<Override PartName={quoteattr('/' + part)} "
                f"ContentType={quoteattr(content_type)}/>"
            )
        lines.append("</Types>")
        return (XML_DECLARATION + "".join(lines)).encode()


class Package:
    """A zip package held in memory, with edits applied on top of it.

    Relationships and content types are edited as objects and written once, by `to_bytes`.
    """

    def __init__(self, data: bytes) -> None:
        self._source = zipfile.ZipFile(io.BytesIO(data))
        self._changes: dict[str, bytes] = {}
        self._rels: dict[str, list[Rel]] = {}
        self.content_types = ContentTypes(self.read(CONTENT_TYPES))

    def names(self) -> set[str]:
        return set(self._source.namelist()) | set(self._changes)

    def exists(self, name: str) -> bool:
        return name in self._changes or name in self._source.namelist()

    def read(self, name: str) -> bytes:
        if name in self._changes:
            return self._changes[name]
        return self._source.read(name)

    def write(self, name: str, data: bytes) -> None:
        self._changes[name] = data

    def rels(self, part: str) -> list[Rel]:
        """The relationships of a part, as a list that is saved when the package is."""
        if part not in self._rels:
            name = rels_name(part)
            self._rels[part] = parse_rels(part, self.read(name)) if self.exists(name) else []
        return self._rels[part]

    def unique_name(self, wanted: str) -> str:
        """``wanted``, or the same name with the next free number in it."""
        taken = self.names()
        if wanted not in taken:
            return wanted
        match = re.fullmatch(r"(.*?)(\d*)(\.[^./]*)?", wanted)
        stem, number, extension = match.groups() if match else (wanted, "", "")
        counter = int(number or 0) + 1
        while f"{stem}{counter}{extension or ''}" in taken:
            counter += 1
        return f"{stem}{counter}{extension or ''}"

    def to_bytes(self) -> bytes:
        for part, rels in self._rels.items():
            name = rels_name(part)
            if rels:
                self._changes[name] = write_rels(rels)
            elif self.exists(name):
                self._changes[name] = b""
        self._changes[CONTENT_TYPES] = self.content_types.to_bytes()
        output = io.BytesIO()
        with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as target:
            for name in self._source.namelist():
                content = self._changes.get(name)
                if content is None:
                    target.writestr(self._source.getinfo(name), self._source.read(name))
                elif content:
                    target.writestr(name, content)
            for name, content in self._changes.items():
                if content and name not in self._source.namelist():
                    target.writestr(name, content)
        return output.getvalue()
