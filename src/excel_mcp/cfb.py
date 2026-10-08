"""Reading and writing Compound File Binary files ([MS-CFB]), the container of vbaProject.bin.

The writer produces version 3 files (512-byte sectors) in the layout Office uses:
small streams live in the mini stream, and the directory entries of each storage
form a balanced tree. It is limited to files whose FAT fits in the header.
"""

import io
import struct
import uuid
from collections.abc import Iterator
from contextlib import contextmanager
from dataclasses import dataclass

import olefile

from excel_mcp.errors import WorkbookError

_SECTOR_SIZE = 512
_MINI_SECTOR_SIZE = 64
_MINI_STREAM_CUTOFF = 4096
_ENTRIES_PER_SECTOR = _SECTOR_SIZE // 4
_MAX_FAT_SECTORS = 109
_DIRECTORY_ENTRY_SIZE = 128

_FREE = 0xFFFFFFFF
_END_OF_CHAIN = 0xFFFFFFFE
_FAT_SECTOR = 0xFFFFFFFD
_NO_STREAM = 0xFFFFFFFF

_STORAGE = 1
_STREAM = 2
_ROOT = 5
_RED = 0
_BLACK = 1


@dataclass(frozen=True)
class Entry:
    """A storage (``data`` is None) or a stream, addressed by its ``/``-separated path."""

    path: str
    data: bytes | None = None
    clsid: str = ""


@contextmanager
def reading(content: bytes) -> Iterator[olefile.OleFileIO]:
    """Open a compound file for reading, reporting damaged files as WorkbookError."""
    if not olefile.isOleFile(io.BytesIO(content)) or not _has_standard_sectors(content):
        raise WorkbookError("The VBA project is not a valid OLE compound file.")
    try:
        with olefile.OleFileIO(io.BytesIO(content)) as ole:
            yield ole
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


def read_entries(content: bytes) -> list[Entry]:
    with reading(content) as ole:
        entries = []
        for parts in ole.listdir(streams=True, storages=True):
            path = "/".join(parts)
            is_stream = ole.get_type(path) == olefile.STGTY_STREAM
            data = ole.openstream(path).read() if is_stream else None
            entries.append(Entry(path, data, ole.getclsid(path)))
        return entries


def write(entries: list[Entry]) -> bytes:
    """Build a compound file; storages that only appear in a path are created implicitly."""
    nodes = _build_nodes(entries)
    big = [node for node in nodes if _is_big(node)]
    small = [node for node in nodes if node.data is not None and not _is_big(node)]

    mini_stream = bytearray()
    mini_fat: list[int] = []
    for node in small:
        node.start = len(mini_fat) if node.data else _END_OF_CHAIN
        mini_stream += _padded(node.data or b"", _MINI_SECTOR_SIZE)
        _append_chain(mini_fat, -(-len(node.data or b"") // _MINI_SECTOR_SIZE))

    fat: list[int] = []
    sectors = bytearray()

    def add_sectors(data: bytes) -> int:
        """Append data as a chain of sectors and return the first sector's number."""
        count = -(-len(data) // _SECTOR_SIZE)
        if count == 0:
            return _END_OF_CHAIN
        first = len(fat)
        _append_chain(fat, count, first)
        sectors.extend(_padded(data, _SECTOR_SIZE))
        return first

    root = nodes[0]
    root.start = add_sectors(bytes(mini_stream))
    root.size = len(mini_stream)
    for node in big:
        node.start = add_sectors(node.data or b"")
    directory_start = add_sectors(_directory(nodes))
    mini_fat_start = add_sectors(_pack_table(mini_fat))

    fat_count = _fat_sector_count(len(fat))
    if fat_count > _MAX_FAT_SECTORS:
        raise WorkbookError("The VBA project is too large to write.")
    fat_start = len(fat)
    fat += [_FAT_SECTOR] * fat_count

    header = _header(fat_count, fat_start, directory_start, mini_fat_start, len(mini_fat))
    return header + bytes(sectors) + _pack_table(fat)


def _fat_sector_count(used: int) -> int:
    count = 0
    while count * _ENTRIES_PER_SECTOR < used + count:
        count += 1
    return count


class _Node:
    def __init__(self, name: str, kind: int, data: bytes | None, clsid: str) -> None:
        self.name = name
        self.kind = kind
        self.data = data
        self.clsid = clsid
        self.size = len(data) if data is not None else 0
        self.start = _END_OF_CHAIN
        self.color = _BLACK
        self.left = self.right = self.child = _NO_STREAM
        self.children: list[_Node] = []


def _is_big(node: _Node) -> bool:
    return node.data is not None and len(node.data) >= _MINI_STREAM_CUTOFF


def _build_nodes(entries: list[Entry]) -> list[_Node]:
    """Number the nodes, root first, and link every storage's children into a tree."""
    root = _Node("Root Entry", _ROOT, None, "")
    by_path: dict[str, _Node] = {"": root}
    for entry in sorted(entries, key=lambda item: item.path):
        parent_path, _, name = entry.path.rpartition("/")
        _storage(by_path, parent_path)
        node = _Node(name, _STREAM if entry.data is not None else _STORAGE, entry.data, entry.clsid)
        by_path[entry.path] = node
        by_path[parent_path].children.append(node)
    nodes = [root]
    pending = [root]
    while pending:
        storage = pending.pop(0)
        storage.children.sort(key=lambda node: (len(node.name), node.name.upper()))
        _link(storage, storage.children, nodes)
        pending += [node for node in storage.children if node.kind == _STORAGE]
    return nodes


def _storage(by_path: dict[str, _Node], path: str) -> None:
    if path in by_path:
        return
    parent_path, _, name = path.rpartition("/")
    _storage(by_path, parent_path)
    node = _Node(name, _STORAGE, None, "")
    by_path[path] = node
    by_path[parent_path].children.append(node)


def _link(storage: _Node, children: list[_Node], nodes: list[_Node]) -> None:
    """Number the children and make them a balanced binary tree hanging off the storage."""
    first = len(nodes)
    nodes.extend(children)
    ids = {id(node): first + index for index, node in enumerate(children)}
    if not children:
        return

    def build(low: int, high: int, depth: int) -> int:
        middle = (low + high) // 2
        node = children[middle]
        node.color = _RED if depth == last_depth and not perfect else _BLACK
        if low < middle:
            node.left = build(low, middle, depth + 1)
        if middle + 1 < high:
            node.right = build(middle + 1, high, depth + 1)
        return ids[id(node)]

    last_depth = len(children).bit_length() - 1
    perfect = (len(children) + 1) & len(children) == 0
    storage.child = build(0, len(children), 0)


def _directory(nodes: list[_Node]) -> bytes:
    out = bytearray()
    for node in nodes:
        name = node.name.encode("utf-16-le")
        clsid = uuid.UUID(node.clsid).bytes_le if node.clsid else bytes(16)
        out += name.ljust(64, b"\0")
        out += struct.pack("<HBB", len(name) + 2, node.kind, node.color)
        out += struct.pack("<III", node.left, node.right, node.child)
        out += clsid + bytes(4 + 16)
        out += struct.pack("<IQ", node.start, node.size)
    unused = struct.pack("<III", _NO_STREAM, _NO_STREAM, _NO_STREAM)
    padding = (-len(nodes)) % (_SECTOR_SIZE // _DIRECTORY_ENTRY_SIZE)
    out += (bytes(68) + unused + bytes(48)) * padding
    return bytes(out)


def _header(
    fat_count: int, fat_start: int, directory_start: int, mini_fat_start: int, mini_fat_len: int
) -> bytes:
    difat = [fat_start + index for index in range(fat_count)]
    difat += [_FREE] * (_MAX_FAT_SECTORS - fat_count)
    mini_fat_sectors = -(-mini_fat_len // _ENTRIES_PER_SECTOR)
    return (
        bytes.fromhex("D0CF11E0A1B11AE1")
        + bytes(16)
        + struct.pack("<HHHHH", 0x3E, 3, 0xFFFE, 9, 6)
        + bytes(6)
        + struct.pack("<IIIII", 0, fat_count, directory_start, 0, _MINI_STREAM_CUTOFF)
        + struct.pack("<IIII", mini_fat_start, mini_fat_sectors, _END_OF_CHAIN, 0)
        + struct.pack(f"<{_MAX_FAT_SECTORS}I", *difat)
    )


def _append_chain(table: list[int], count: int, first: int | None = None) -> None:
    """Append the table entries of a chain of ``count`` sectors."""
    start = len(table) if first is None else first
    table.extend(range(start + 1, start + count))
    if count:
        table.append(_END_OF_CHAIN)


def _pack_table(values: list[int]) -> bytes:
    """Pack allocation table entries, filling the last sector with free entries."""
    values = values + [_FREE] * ((-len(values)) % _ENTRIES_PER_SECTOR)
    return struct.pack(f"<{len(values)}I", *values)


def _padded(data: bytes, unit: int) -> bytes:
    return data + bytes((-len(data)) % unit)
