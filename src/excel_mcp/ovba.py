"""Reading VBA source code from a vbaProject.bin file, following [MS-OVBA].

A VBA project is an OLE compound file. Its ``VBA/dir`` stream lists the
modules; each module's source sits in its own stream after a p-code block,
compressed with the MS-OVBA run-length scheme. This module decompresses the
streams and parses just the records needed to find the source code.
"""

import codecs
import io
from dataclasses import dataclass

import olefile

from excel_mcp.errors import WorkbookError

MAX_DECOMPRESSED_BYTES = 16 * 1024 * 1024

_CHUNK_SIZE = 4096
_PROJECT_CODE_PAGE = 0x0003
_PROJECT_VERSION = 0x0009
_MODULE_NAME = 0x0019
_MODULE_NAME_UNICODE = 0x0047
_MODULE_STREAM_NAME = 0x001A
_MODULE_STREAM_NAME_UNICODE = 0x0032
_MODULE_OFFSET = 0x0031
_MODULE_TYPE_PROCEDURAL = 0x0021
_MODULE_END = 0x002B


@dataclass(frozen=True)
class RawModule:
    name: str
    procedural: bool
    source: str


def read_project(content: bytes) -> tuple[list[RawModule], str]:
    """Return the project's modules and the text of its PROJECT stream."""
    if not olefile.isOleFile(io.BytesIO(content)) or not _has_standard_sectors(content):
        raise WorkbookError("The VBA project is not a valid OLE compound file.")
    try:
        return _read_modules(content)
    except (OSError, ValueError) as error:
        # olefile reports damaged containers with these exception types.
        raise WorkbookError(f"The VBA project could not be read ({error}).") from None


def _has_standard_sectors(content: bytes) -> bool:
    """Office writes 512- or 4096-byte sectors and 64-byte mini sectors.

    Checked before olefile parses the file, because olefile computes 2**shift
    from these header fields without bounding them first.
    """
    sector_shift = int.from_bytes(content[30:32], "little")
    mini_sector_shift = int.from_bytes(content[32:34], "little")
    return sector_shift in (9, 12) and mini_sector_shift == 6


def _read_modules(content: bytes) -> tuple[list[RawModule], str]:
    with olefile.OleFileIO(io.BytesIO(content)) as ole:
        directory = decompress(_read_stream(ole, "VBA/dir"))
        code_page, entries = _parse_directory(directory)
        encoding = _encoding(code_page)
        modules = [
            RawModule(
                name=name,
                procedural=procedural,
                source=decompress(_read_stream(ole, f"VBA/{stream}")[offset:]).decode(
                    encoding, errors="replace"
                ),
            )
            for name, stream, offset, procedural in entries
        ]
        project_text = _read_stream(ole, "PROJECT").decode(encoding, errors="replace")
    return modules, project_text


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


def _parse_directory(data: bytes) -> tuple[int, list[tuple[str, str, int, bool]]]:
    """Read the code page and, per module, its name, stream name, source offset and type."""
    code_page = 1252
    modules: list[tuple[str, str, int, bool]] = []
    current: dict[int, bytes] = {}
    position = 0
    while position + 6 <= len(data):
        record_id = int.from_bytes(data[position : position + 2], "little")
        size = int.from_bytes(data[position + 2 : position + 6], "little")
        # PROJECTVERSION declares a size of 4 but carries 6 bytes of data.
        if record_id == _PROJECT_VERSION:
            size = 6
        value = data[position + 6 : position + 6 + size]
        position += 6 + size

        if record_id == _PROJECT_CODE_PAGE:
            code_page = _integer(value, 2)
        elif record_id in _MODULE_RECORDS:
            current[record_id] = value
        elif record_id == _MODULE_TYPE_PROCEDURAL:
            current[record_id] = b""
        elif record_id == _MODULE_END:
            modules.append(_module_entry(current, code_page))
            current = {}
    return code_page, modules


_MODULE_RECORDS = frozenset(
    {
        _MODULE_NAME,
        _MODULE_NAME_UNICODE,
        _MODULE_STREAM_NAME,
        _MODULE_STREAM_NAME_UNICODE,
        _MODULE_OFFSET,
    }
)


def _module_entry(records: dict[int, bytes], code_page: int) -> tuple[str, str, int, bool]:
    if _MODULE_OFFSET not in records or _MODULE_NAME not in records:
        raise WorkbookError("The VBA project directory is missing module information.")
    name = _text(records, _MODULE_NAME_UNICODE, _MODULE_NAME, code_page)
    stream = _text(records, _MODULE_STREAM_NAME_UNICODE, _MODULE_STREAM_NAME, code_page)
    offset = _integer(records[_MODULE_OFFSET], 4)
    return name, stream, offset, _MODULE_TYPE_PROCEDURAL in records


def _text(records: dict[int, bytes], unicode_id: int, mbcs_id: int, code_page: int) -> str:
    """Names are stored twice; prefer the UTF-16 copy, which newer Office versions write."""
    if records.get(unicode_id):
        return records[unicode_id].decode("utf-16-le", errors="replace")
    return records.get(mbcs_id, b"").decode(_encoding(code_page), errors="replace")


def _integer(value: bytes, size: int) -> int:
    if len(value) != size:
        raise WorkbookError("The VBA project directory contains a malformed record.")
    return int.from_bytes(value, "little")


def _encoding(code_page: int) -> str:
    encoding = f"cp{code_page}"
    try:
        codecs.lookup(encoding)
    except LookupError:
        raise WorkbookError(
            f"The VBA project uses an unsupported code page ({code_page})."
        ) from None
    return encoding


def _read_stream(ole: olefile.OleFileIO, name: str) -> bytes:
    if not ole.exists(name):
        raise WorkbookError(f"The VBA project has no {name} stream.")
    return ole.openstream(name).read()
