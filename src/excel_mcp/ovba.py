"""Reading VBA source code from a vbaProject.bin file, following [MS-OVBA].

A VBA project is an OLE compound file. Its ``VBA/dir`` stream lists the
modules; each module's source sits in its own stream after a p-code block,
compressed with the MS-OVBA run-length scheme. This module decompresses the
streams and parses just the records needed to find the source code.
"""

import codecs
from collections.abc import Iterator
from dataclasses import dataclass
from typing import Literal

import olefile

from excel_mcp import cfb
from excel_mcp.errors import WorkbookError

MAX_DECOMPRESSED_BYTES = 16 * 1024 * 1024

ModuleKind = Literal["standard", "class", "document", "form"]
Record = tuple[int, bytes]

_CHUNK_SIZE = 4096
PROJECT_CODE_PAGE = 0x0003
PROJECT_VERSION = 0x0009
PROJECT_MODULES = 0x000F
MODULE_NAME = 0x0019
MODULE_OFFSET = 0x0031
MODULE_TYPE_PROCEDURAL = 0x0021
MODULE_END = 0x002B
_MODULE_NAME_UNICODE = 0x0047
_MODULE_STREAM_NAME = 0x001A
_MODULE_STREAM_NAME_UNICODE = 0x0032


@dataclass(frozen=True)
class RawModule:
    name: str
    stream: str
    kind: ModuleKind
    source: str
    records: list[Record]


@dataclass(frozen=True)
class ProjectData:
    code_page: int
    header: list[Record]
    modules: list[RawModule]
    text: str


def read_project(content: bytes) -> ProjectData:
    """Read the project's modules, the directory records before them and the PROJECT stream."""
    with cfb.reading(content) as ole:
        directory = decompress(_read_stream(ole, "VBA/dir"))
        code_page, header, entries = _parse_directory(directory)
        text = _read_stream(ole, "PROJECT").decode(encoding(code_page), errors="replace")
        kinds = module_kinds(text)
        modules = [
            RawModule(
                name=name,
                stream=stream,
                kind="standard" if procedural else kinds.get(name.casefold(), "class"),
                source=decompress(_read_stream(ole, f"VBA/{stream}")[offset:])
                .decode(encoding(code_page), errors="replace")
                .rstrip("\0"),
                records=records,
            )
            for name, stream, offset, procedural, records in entries
        ]
    return ProjectData(code_page, header, modules, text)


def module_kinds(project_text: str) -> dict[str, ModuleKind]:
    """The PROJECT stream tells document modules and forms apart from classes."""
    kinds: dict[str, ModuleKind] = {}
    for line in project_text.splitlines():
        key, _, value = line.partition("=")
        if key == "Document":
            kinds[value.split("/")[0].casefold()] = "document"
        elif key == "BaseClass":
            kinds[value.casefold()] = "form"
    return kinds


def decompress(data: bytes) -> bytes:
    """Decompress an MS-OVBA compressed container (section 2.4.1)."""
    if not data or data[0] != 0x01:
        raise WorkbookError("The VBA project contains a stream that is not compressed.")
    output = bytearray()
    position = 1
    while position < len(data):
        if position + 2 > len(data):
            raise WorkbookError("The VBA project contains a truncated stream.")
        header = int.from_bytes(data[position : position + 2], "little")
        chunk_end = min(position + (header & 0x0FFF) + 3, len(data))
        position += 2
        if header & 0x8000:
            _decompress_chunk(data, position, chunk_end, output)
            position = chunk_end
        else:
            output += data[position : position + _CHUNK_SIZE]
            position += _CHUNK_SIZE
        if len(output) > MAX_DECOMPRESSED_BYTES:
            raise WorkbookError("The VBA project is too large to read.")
    return bytes(output)


def _decompress_chunk(data: bytes, position: int, end: int, output: bytearray) -> None:
    chunk_start = len(output)
    while position < end:
        flags = data[position]
        position += 1
        for bit in range(8):
            if position >= end:
                return
            if not flags & (1 << bit):
                output.append(data[position])
                position += 1
                continue
            if position + 2 > end:
                raise WorkbookError("The VBA project contains a corrupt stream.")
            token = int.from_bytes(data[position : position + 2], "little")
            position += 2
            decompressed = len(output) - chunk_start
            bit_count = max((decompressed - 1).bit_length(), 4)
            offset = (token >> (16 - bit_count)) + 1
            length = (token & (0xFFFF >> bit_count)) + 3
            if offset > decompressed:
                raise WorkbookError("The VBA project contains a corrupt stream.")
            for _ in range(length):
                output.append(output[-offset])


def records(data: bytes) -> Iterator[Record]:
    """Walk the (id, value) records of the ``dir`` stream."""
    position = 0
    while position + 6 <= len(data):
        record_id = int.from_bytes(data[position : position + 2], "little")
        size = int.from_bytes(data[position + 2 : position + 6], "little")
        # PROJECTVERSION declares a size of 4 but carries 6 bytes of data.
        if record_id == PROJECT_VERSION:
            size = 6
        yield record_id, data[position + 6 : position + 6 + size]
        position += 6 + size


_ModuleEntry = tuple[str, str, int, bool, list[Record]]


def _parse_directory(data: bytes) -> tuple[int, list[Record], list[_ModuleEntry]]:
    """Split the directory into the records before the modules and the modules."""
    header: list[Record] = []
    modules: list[_ModuleEntry] = []
    code_page = 1252
    in_modules = False
    current: list[Record] | None = None
    for record_id, value in records(data):
        if not in_modules and record_id != PROJECT_MODULES:
            header.append((record_id, value))
            if record_id == PROJECT_CODE_PAGE:
                code_page = _integer(value, 2)
        elif record_id == PROJECT_MODULES:
            in_modules = True
        elif record_id == MODULE_NAME:
            current = [(record_id, value)]
        elif current is not None:
            current.append((record_id, value))
            if record_id == MODULE_END:
                modules.append(_module_entry(current, code_page))
                current = None
    return code_page, header, modules


def _module_entry(module_records: list[Record], code_page: int) -> _ModuleEntry:
    by_id = dict(module_records)
    if MODULE_OFFSET not in by_id:
        raise WorkbookError("The VBA project directory is missing module information.")
    name = _text(by_id, _MODULE_NAME_UNICODE, MODULE_NAME, code_page)
    stream = _text(by_id, _MODULE_STREAM_NAME_UNICODE, _MODULE_STREAM_NAME, code_page)
    offset = _integer(by_id[MODULE_OFFSET], 4)
    return name, stream, offset, MODULE_TYPE_PROCEDURAL in by_id, module_records


def _text(by_id: dict[int, bytes], unicode_id: int, mbcs_id: int, code_page: int) -> str:
    """Names are stored twice; prefer the UTF-16 copy, which newer Office versions write."""
    if by_id.get(unicode_id):
        return by_id[unicode_id].decode("utf-16-le", errors="replace")
    return by_id.get(mbcs_id, b"").decode(encoding(code_page), errors="replace")


def _integer(value: bytes, size: int) -> int:
    if len(value) != size:
        raise WorkbookError("The VBA project directory contains a malformed record.")
    return int.from_bytes(value, "little")


def encoding(code_page: int) -> str:
    name = f"cp{code_page}"
    try:
        codecs.lookup(name)
    except LookupError:
        raise WorkbookError(
            f"The VBA project uses an unsupported code page ({code_page})."
        ) from None
    return name


def _read_stream(ole: olefile.OleFileIO, name: str) -> bytes:
    if not ole.exists(name):
        raise WorkbookError(f"The VBA project has no {name} stream.")
    return ole.openstream(name).read()
