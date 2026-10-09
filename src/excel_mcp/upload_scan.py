"""Streaming safety scan of uploaded workbook packages.

An uploaded file is stored unchanged, so everything in its zip is checked here
rather than only what openpyxl would load: every entry that holds XML is found
by its content (not its name), parsed in constant memory, and searched for
formulas. Packages are refused when their entries could escape a folder or
shadow each other, when they expand beyond the limits, or when they declare
data connections, query tables or external links, which fetch remote data.
"""

import io
import re
import zipfile
import zlib
from collections.abc import Iterator
from typing import IO, NamedTuple
from xml.etree import ElementTree
from xml.sax.saxutils import unescape

from excel_mcp.config import Limits
from excel_mcp.errors import InvalidArgumentError, LimitExceededError, UnsafeFormulaError
from excel_mcp.links import is_allowed_address

# These elements hold formulas: cell formulas and chart or sparkline references
# ("f"), rule formulas, table column formulas and defined names.
_FORMULA_ELEMENTS = frozenset(
    {
        "f",
        "formula",
        "formula1",
        "formula2",
        "calculatedColumnFormula",
        "totalsRowFormula",
        "definedName",
    }
)
_PIVOT_FORMULA_ELEMENTS = frozenset({"cacheField", "calculatedItem"})
_CONTROL_LINKS = ("fmlaLink", "fmlaRange", "fmlaGroup", "fmlaTxbx")
_CONTENT_TYPES = "[Content_Types].xml"
# Namespaces whose formulas Excel evaluates: SpreadsheetML (and its 2009-2016 extensions: x14,
# x15, slicers, timelines), the Excel macro namespace "xm", charts and Excel 2016 charts.
_EXCEL_NAMESPACES = (
    "http://schemas.openxmlformats.org/spreadsheetml/",
    "http://schemas.microsoft.com/office/spreadsheetml/",
    "http://schemas.microsoft.com/office/excel/",
    "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "http://schemas.microsoft.com/office/drawing/",
)
_EXCEL_TYPES = ("spreadsheetml", "drawingml.chart", "vnd.ms-office", "vnd.ms-excel")
_HREF = re.compile(r"""\bhref\s*=\s*(?:"([^"]*)"|'([^']*)')""", re.IGNORECASE)
_MAX_LINK = 4096
_MAX_ENTRIES = 10_000
_MAX_XML_DEPTH = 100
_CHUNK = 64 * 1024
# Small entries are exempt from the ratio limit: tiny parts compress extremely well.
_RATIO_FLOOR = 1024 * 1024
_REMOTE_DATA = ("connections", "querytable", "externallink")
_REMOTE_DATA_ROOTS = frozenset({"connections", "queryTable", "externalLink"})


class Formula(NamedTuple):
    text: str
    # PivotTable formulas name fields and items, so only their safety can be checked.
    pivot: bool = False


def scan_package(content: bytes, limits: Limits, *, check_links: bool = True) -> Iterator[Formula]:
    """Check an uploaded package and yield each formula in it; raises when it is unsafe.

    Nothing is vouched for until the iterator is exhausted.
    """
    try:
        with zipfile.ZipFile(io.BytesIO(content)) as archive:
            infos = archive.infolist()
            _check_entries(infos, limits)
            remaining = limits.max_unpacked_bytes
            types = _Types(check_links)
            # The content types come first: they say which parts Excel reads as its own.
            for info in sorted(infos, key=lambda i: i.filename != _CONTENT_TYPES):
                if info.is_dir():
                    continue
                with archive.open(info) as source:
                    entry = _Entry(source, info, limits, remaining)
                    if info.filename.casefold().endswith(".vml"):
                        _scan_vml(entry, check_links)
                    elif entry.is_xml:
                        yield from _scan_xml(entry, info.filename, types)
                    entry.drain()
                    remaining = entry.remaining
    except (zipfile.BadZipFile, zlib.error, EOFError, NotImplementedError, RuntimeError):
        raise InvalidArgumentError(
            "The uploaded workbook is a damaged or unsupported zip."
        ) from None


def _check_entries(infos: list[zipfile.ZipInfo], limits: Limits) -> None:
    if len(infos) > _MAX_ENTRIES:
        raise LimitExceededError(f"The workbook has more than {_MAX_ENTRIES:,} parts.")
    names: set[str] = set()
    for info in infos:
        name = info.filename
        segments = name.split("/")
        if (
            "\\" in name
            or "\x00" in name
            or name.startswith("/")
            or ".." in segments
            or ":" in segments[0]
        ):
            raise InvalidArgumentError(f"The workbook has a part with an unsafe name: {name!r}.")
        # Part names are case-insensitive, so two spellings are the same part.
        if name.casefold() in names:
            raise InvalidArgumentError(f"The workbook has more than one part named {name!r}.")
        names.add(name.casefold())
        _check_ratio(info.file_size, info.compress_size, limits)
    if sum(info.file_size for info in infos) > limits.max_unpacked_bytes:
        raise LimitExceededError(
            f"The workbook expands to more than {limits.max_unpacked_bytes:,} bytes."
        )


def _check_ratio(size: int, compressed: int, limits: Limits) -> None:
    if size > max(_RATIO_FLOOR, compressed * limits.max_compression_ratio):
        raise LimitExceededError("The workbook is compressed too densely to be a real workbook.")


class _Entry:
    """A zip entry whose limits are enforced as it is decompressed, whatever its header says."""

    def __init__(
        self, source: IO[bytes], info: zipfile.ZipInfo, limits: Limits, remaining: int
    ) -> None:
        self._source = source
        self._compressed = info.compress_size
        self._limits = limits
        self._read = 0
        self.remaining = remaining
        self._head = self._read_counted(64)
        self.is_xml = _looks_like_xml(self._head)

    def read(self, size: int = _CHUNK) -> bytes:
        if self._head:
            head, self._head = self._head, b""
            return head
        return self._read_counted(size)

    def drain(self) -> None:
        while self.read():
            pass

    def _read_counted(self, size: int) -> bytes:
        data = self._source.read(size)
        self._read += len(data)
        self.remaining -= len(data)
        if self.remaining < 0:
            raise LimitExceededError(
                f"The workbook expands to more than {self._limits.max_unpacked_bytes:,} bytes."
            )
        _check_ratio(self._read, self._compressed, self._limits)
        return data


def _looks_like_xml(head: bytes) -> bool:
    if head.startswith((b"\xff\xfe", b"\xfe\xff", b"<\x00", b"\x00<")):
        return True
    return head.removeprefix(b"\xef\xbb\xbf").lstrip().startswith(b"<")


class _Types:
    """The content types a package declares, by part name and by file extension."""

    def __init__(self, check_links: bool) -> None:
        self.check_links = check_links
        self.overrides: dict[str, str] = {}
        self.defaults: dict[str, str] = {}

    def claims_excel(self, part: str) -> bool:
        declared = self.overrides.get(part.casefold()) or self.defaults.get(
            part.rpartition(".")[2].casefold(), ""
        )
        return any(kind in declared.casefold() for kind in _EXCEL_TYPES)


def _reads_formulas(namespace: str, claims_excel: bool) -> bool:
    """Excel evaluates formulas in its own namespaces, and in any part declared as its own."""
    return claims_excel or namespace.startswith(_EXCEL_NAMESPACES)


def _formulas(name: str, element: ElementTree.Element) -> Iterator[Formula]:
    if name in _FORMULA_ELEMENTS and element.text and element.text.strip():
        yield Formula(element.text)
    # Calculated PivotTable fields and items.
    elif name in _PIVOT_FORMULA_ELEMENTS and (value := element.get("formula")):
        yield Formula(value, pivot=True)
    # Form controls name the cells and ranges they are linked to.
    elif name == "formControlPr":
        for attribute in _CONTROL_LINKS:
            if value := element.get(attribute):
                yield Formula(value, pivot=True)
    # Color scale, data bar and icon set thresholds can be formulas too.
    elif name == "cfvo" and (value := element.get("val")):
        yield Formula(value)


def _scan_xml(source: _Entry, part: str, types: _Types) -> Iterator[Formula]:
    stack: list[ElementTree.Element] = []
    claims_excel = types.claims_excel(part)
    try:
        for event, element in ElementTree.iterparse(source, events=("start", "end")):
            namespace, _, name = element.tag.rpartition("}")
            if event == "start":
                if not stack and name in _REMOTE_DATA_ROOTS:
                    _refuse_remote_data()
                stack.append(element)
                if len(stack) > _MAX_XML_DEPTH:
                    raise LimitExceededError(f"XML in {part} is nested too deeply.")
                _check_declaration(name, element, types)
                continue
            stack.pop()
            if _reads_formulas(namespace.removeprefix("{"), claims_excel):
                yield from _formulas(name, element)
            # Freeing each element keeps memory flat however large the part is.
            element.clear()
            if stack:
                stack[-1].remove(element)
    except ElementTree.ParseError:
        # Excel reads legacy drawings (.vml) as lenient HTML, so those may not be well-formed.
        if part.casefold().endswith((".xml", ".rels")):
            raise InvalidArgumentError(
                f"The uploaded workbook part {part} is not valid XML."
            ) from None


def _check_declaration(name: str, element: ElementTree.Element, types: _Types) -> None:
    if name in ("Default", "Override"):
        declared = element.get("ContentType", "")
        if name == "Override":
            types.overrides[element.get("PartName", "").removeprefix("/").casefold()] = declared
        else:
            types.defaults[element.get("Extension", "").casefold()] = declared
    elif name == "Relationship":
        declared = element.get("Type", "")
        external = element.get("TargetMode", "").casefold() == "external"
        # Only a hyperlink may point outside the file; it is followed when the user clicks it.
        if external and not declared.casefold().endswith("/hyperlink"):
            _refuse_remote_data()
        if external and types.check_links:
            _check_link(element.get("Target", ""))
    elif name == "oleLink" or (name == "oleObject" and element.get("link")):
        _refuse_remote_data()
        return
    else:
        return
    if any(kind in declared.casefold() for kind in _REMOTE_DATA):
        _refuse_remote_data()


def _scan_vml(entry: _Entry, check_links: bool) -> None:
    """Check the links of a legacy drawing. Excel reads these as lenient HTML, so the text is
    searched instead of parsed, which a broken tag cannot get around."""
    tail = ""
    while chunk := entry.read():
        text = tail + chunk.decode("utf-8", "ignore")
        for found in _HREF.finditer(text) if check_links else ():
            _check_link(unescape(found[1] if found[1] is not None else found[2]))
        tail = text[-_MAX_LINK:]


def _check_link(target: str) -> None:
    """A hyperlink may lead to a place in the workbook or to a plain web or mail address."""
    if not target.startswith("#") and not is_allowed_address(target):
        raise UnsafeFormulaError(
            f"The workbook has a hyperlink to {target[:200]!r}, which is not allowed: only "
            "http://, https:// and mailto: links without credentials, and places in the "
            "workbook, can be uploaded."
        )


def _refuse_remote_data() -> None:
    raise UnsafeFormulaError(
        "Workbooks with data connections, query tables, linked objects or links to other "
        "workbooks or files cannot be uploaded, because they reach outside the file."
    )
