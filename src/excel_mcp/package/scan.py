"""Splitting XML documents into their top-level elements without re-serializing them.

Preserved content is carried as the exact text Excel wrote. Parsing and serializing it
again would drop namespace declarations that only ``mc:Ignorable`` refers to, which makes
Excel reject the file.
"""

import re
from dataclasses import dataclass
from xml.parsers import expat

from excel_mcp.errors import WorkbookError

REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_START_TAG = re.compile(rb"<[^\s>/]+(?:\s+[^\s=>/]+\s*=\s*(?:\"[^\"]*\"|'[^']*'))*\s*/?>")
_ROOT = re.compile(rb"<[A-Za-z_][^\s>/]*")
_NAME = re.compile(r"<([^\s>/]+)")
_DECLARED = re.compile(r"xmlns(?::([\w.-]+))?\s*=\s*\"([^\"]*)\"")
_ATTRIBUTE = re.compile(r"""([\w:.-]+)\s*=\s*(?:"([^"]*)"|'([^']*)')""")
_REL_ATTRIBUTE = re.compile(r"""(?<=[\s"'])([\w.-]+):([\w.-]+)(\s*=\s*)(["'])([^"']*)\4""")


@dataclass(frozen=True)
class Child:
    """A top-level element: its namespace, names and the byte span it occupies."""

    uri: str
    local: str
    qname: str
    start: int
    end: int


@dataclass(frozen=True)
class Scan:
    data: bytes
    uri: str
    local: str
    namespaces: dict[str, str]
    root_tag: tuple[int, int]
    close_tag: int
    children: list[Child]

    def raw(self, child: Child) -> str:
        return self.data[child.start : child.end].decode("utf-8")


def scan(data: bytes, skip: str | None = None) -> Scan:
    """Locate the root element and each of its children.

    ``skip`` names a child (``sheetData``) to jump over instead of parsing: it can be
    hundreds of megabytes, and it is neither listed nor needed.
    """
    first = _ROOT.search(data)
    root = _START_TAG.match(data, first.start()) if first else None
    closing = data.rfind(f"</{skip}>".encode()) if skip else -1
    if root is None or closing < 0:
        return _scan(data, data, 0)
    tail = closing + len(skip or "") + 3
    return _scan(data, data[: root.end()] + data[tail:], tail - root.end())


def _scan(data: bytes, document: bytes, shift: int) -> Scan:
    parser = expat.ParserCreate(namespace_separator=" ")
    namespaces: dict[str, str] = {}
    children: list[Child] = []
    state = {"depth": 0, "root": "", "start": 0, "root_tag": (0, 0), "close": 0}

    def declaration(prefix: str | None, uri: str | None) -> None:
        if state["depth"] == 0:
            namespaces[prefix or ""] = uri or ""

    def start(name: str, _attributes: dict[str, str]) -> None:
        index = parser.CurrentByteIndex
        if state["depth"] == 0:
            tag = _START_TAG.match(document, index)
            state.update(root=name, root_tag=(index, tag.end() if tag else index))
        elif state["depth"] == 1:
            state["start"] = index
        state["depth"] += 1

    def end(name: str) -> None:
        index = parser.CurrentByteIndex
        state["depth"] -= 1
        if state["depth"] == 1:
            children.append(_child(document, name, state["start"], index, shift))
        elif state["depth"] == 0:
            state["close"] = index + shift

    def reject_dtd(*_: object) -> None:
        raise WorkbookError("XML documents with a DOCTYPE declaration are not supported.")

    parser.StartNamespaceDeclHandler = declaration
    parser.StartElementHandler = start
    parser.EndElementHandler = end
    parser.StartDoctypeDeclHandler = reject_dtd
    try:
        parser.Parse(document, True)
    except expat.ExpatError as error:
        raise WorkbookError(f"A part of the workbook is not well-formed XML ({error}).") from None
    uri, _, local = state["root"].rpartition(" ")
    return Scan(data, uri, local, namespaces, state["root_tag"], state["close"], children)


def _child(document: bytes, name: str, start: int, end_index: int, shift: int) -> Child:
    uri, _, local = name.rpartition(" ")
    tag = _START_TAG.match(document, start)
    if tag is None:
        raise WorkbookError("A part of the workbook is not well-formed XML.")
    # For <a/> expat reports the end of the tag; otherwise the start of </a>.
    end = tag.end() if tag.group(0).endswith(b"/>") else document.index(b">", end_index) + 1
    qname = _NAME.match(document[start : start + 200].decode("utf-8", "ignore"))
    return Child(uri, local, qname.group(1) if qname else local, start + shift, end + shift)


def scan_fragment(xml: str, namespaces: dict[str, str]) -> Scan:
    """Scan a preserved element on its own: its children are the ones listed."""
    return scan(with_namespaces(xml, namespaces).encode())


def used_prefixes(xml: str) -> set[str]:
    """Prefixes that element and attribute names in the text carry."""
    return set(re.findall(r"[<\s/]([A-Za-z_][\w.-]*):[\w.-]+", xml))


def with_namespaces(xml: str, namespaces: dict[str, str]) -> str:
    """Declare, on the element's own start tag, the prefixes it uses but does not declare."""
    declared = {prefix for prefix, _ in _DECLARED.findall(xml)}
    name = _NAME.match(xml)
    if name is None:
        raise WorkbookError("Preserved content is not an XML element.")
    required = " ".join(re.findall(r'Requires="([^"]*)"', xml)).split()
    wanted = used_prefixes(xml) | set(required)
    missing = sorted(p for p in wanted if p in namespaces and p not in declared)
    declarations = "".join(f' xmlns{":" + p if p else ""}="{namespaces[p]}"' for p in missing)
    return xml.replace(name[0], name[0] + declarations, 1) if declarations else xml


def relationship_ids(xml: str, scope: dict[str, str]) -> set[str]:
    """Values of the attributes in the relationships namespace (``r:id``, ``r:embed``, ...)."""
    prefixes = _relationship_prefixes(xml, scope)
    return {m[5] for m in _REL_ATTRIBUTE.finditer(xml) if _names_relationship(m, prefixes)}


def remap_relationship_ids(xml: str, mapping: dict[str, str], scope: dict[str, str]) -> str:
    prefixes = _relationship_prefixes(xml, scope)

    def replace(match: re.Match[str]) -> str:
        if not _names_relationship(match, prefixes) or match[5] not in mapping:
            return match[0]
        return f"{match[1]}:{match[2]}{match[3]}{match[4]}{mapping[match[5]]}{match[4]}"

    return _REL_ATTRIBUTE.sub(replace, xml)


def _names_relationship(match: re.Match[str], prefixes: set[str]) -> bool:
    """Whether an attribute holds a relationship id: ``r:id``, or ``o:relid`` of VML."""
    return match[1] in prefixes or match[2] == "relid"


def _relationship_prefixes(xml: str, scope: dict[str, str]) -> set[str]:
    known = {**scope, **dict(_DECLARED.findall(xml))}
    return {prefix for prefix, uri in known.items() if uri == REL_NS and prefix}


def attributes(tag: str) -> dict[str, str]:
    """The attributes of a start tag, unescaped."""
    return {m[1]: unescape(m[2] if m[2] is not None else m[3]) for m in _ATTRIBUTE.finditer(tag)}


def unescape(value: str) -> str:
    return (
        value.replace("&lt;", "<")
        .replace("&gt;", ">")
        .replace("&quot;", '"')
        .replace("&apos;", "'")
        .replace("&amp;", "&")
    )


def tags(document: Scan, name: str) -> list[dict[str, str]]:
    """The attributes of each ``<name ...>`` start tag in a small document."""
    return [attributes(t) for t in re.findall(rf"<{name}\b[^>]*>", document.data.decode("utf-8"))]


def rel_id(values: dict[str, str], document: Scan) -> str:
    """The relationship id (``r:id``) among the attributes of a tag of ``document``."""
    prefix = next((p for p, uri in document.namespaces.items() if uri == REL_NS), None)
    return values.get(f"{prefix}:id", "") if prefix else ""
