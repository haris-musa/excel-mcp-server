"""Streaming safety scan of uploaded workbook packages.

An uploaded file is stored unchanged, so everything in its zip is checked here
rather than only what openpyxl would load: every entry that holds XML is found
by its content (not its name), parsed in constant memory, and searched for
formulas. Packages are refused when their entries could escape a folder or
shadow each other, when they expand beyond the limits, or when they declare
data connections, query tables or external links, which fetch remote data.
"""

import io
import zipfile
import zlib
from collections.abc import Iterator
from typing import IO
from xml.etree import ElementTree

from excel_mcp.config import Limits
from excel_mcp.errors import InvalidArgumentError, LimitExceededError, UnsafeFormulaError

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
_MAX_ENTRIES = 10_000
_MAX_XML_DEPTH = 100
_CHUNK = 64 * 1024
# Small entries are exempt from the ratio limit: tiny parts compress extremely well.
_RATIO_FLOOR = 1024 * 1024
_REMOTE_DATA = ("connections", "querytable", "externallink")
_REMOTE_DATA_ROOTS = frozenset({"connections", "queryTable", "externalLink"})


def scan_package(content: bytes, limits: Limits) -> Iterator[str]:
    """Check an uploaded package and yield each formula in it; raises when it is unsafe.

    Nothing is vouched for until the iterator is exhausted.
    """
    try:
        with zipfile.ZipFile(io.BytesIO(content)) as archive:
            infos = archive.infolist()
            _check_entries(infos, limits)
            remaining = limits.max_unpacked_bytes
            for info in infos:
                if info.is_dir():
                    continue
                with archive.open(info) as source:
                    entry = _Entry(source, info, limits, remaining)
                    if entry.is_xml:
                        yield from _scan_xml(entry, info.filename)
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


def _scan_xml(source: _Entry, part: str) -> Iterator[str]:
    stack: list[ElementTree.Element] = []
    try:
        for event, element in ElementTree.iterparse(source, events=("start", "end")):
            name = element.tag.rpartition("}")[2]
            if event == "start":
                if not stack and name in _REMOTE_DATA_ROOTS:
                    _refuse_remote_data()
                stack.append(element)
                if len(stack) > _MAX_XML_DEPTH:
                    raise LimitExceededError(f"XML in {part} is nested too deeply.")
                _check_declaration(name, element)
                continue
            stack.pop()
            if name in _FORMULA_ELEMENTS and element.text and element.text.strip():
                yield element.text
            # Color scale, data bar and icon set thresholds can be formulas too.
            elif name == "cfvo" and (value := element.get("val")):
                yield value
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


def _check_declaration(name: str, element: ElementTree.Element) -> None:
    if name in ("Default", "Override"):
        declared = element.get("ContentType", "")
    elif name == "Relationship":
        declared = element.get("Type", "")
    else:
        return
    if any(kind in declared.casefold() for kind in _REMOTE_DATA):
        _refuse_remote_data()


def _refuse_remote_data() -> None:
    raise UnsafeFormulaError(
        "Workbooks with data connections, query tables or links to other workbooks "
        "cannot be uploaded, because they fetch data from elsewhere."
    )
