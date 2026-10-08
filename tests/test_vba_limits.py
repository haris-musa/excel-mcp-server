"""Untrusted vbaProject.bin files fail fast instead of exhausting time or memory."""

import contextlib
import random
import struct
import time
from collections.abc import Callable

import pytest
from openpyxl import Workbook

from excel_mcp import cfb, macros, ovba
from excel_mcp.errors import ExcelMCPError, WorkbookError
from excel_mcp.ovba_compress import compress
from excel_mcp.ovba_write import Project


def _project(code: str = "Sub A()\nEnd Sub") -> Project:
    project = macros.new_project(Workbook())
    project.set_code("Module1", code, "standard")
    return project


def _with_directory_of(project: Project, extra_modules: int) -> bytes:
    """A project file whose directory lists the last module ``extra_modules`` more times."""
    entries = cfb.read_entries(project.to_bytes())
    project.modules += [project.modules[-1]] * extra_modules
    directory = cfb.Entry("VBA/dir", compress(project._directory()))  # pyright: ignore[reportPrivateUsage]
    return cfb.write([directory if e.path == "VBA/dir" else e for e in entries])


def test_many_modules_are_refused() -> None:
    content = _with_directory_of(_project(), ovba.MAX_MODULES)
    with pytest.raises(WorkbookError, match="more than"):
        ovba.read_project(content)


def test_modules_sharing_one_large_stream_cannot_multiply_the_work() -> None:
    megabyte = 1024 * 1024
    content = _with_directory_of(_project("a" * (4 * megabyte)), 5)
    started = time.perf_counter()
    with pytest.raises(WorkbookError, match="too large"):
        ovba.read_project(content)
    assert time.perf_counter() - started < 20


def test_a_large_directory_is_refused() -> None:
    data = compress(bytes(ovba.MAX_DIRECTORY_BYTES + 10))
    with pytest.raises(WorkbookError, match="too large"):
        ovba.decompress(data, ovba.MAX_DIRECTORY_BYTES)


def _fat_offset(data: bytes | bytearray) -> int:
    return 512 + struct.unpack_from("<I", data, 76)[0] * 512


def _fat_self_loops(data: bytearray) -> None:
    offset = _fat_offset(data)
    for index in range(128):
        if struct.unpack_from("<I", data, offset + 4 * index)[0] < 0xFFFFFFF0:
            struct.pack_into("<I", data, offset + 4 * index, index)


def _directory_loop(data: bytearray) -> None:
    start = struct.unpack_from("<I", data, 48)[0]
    struct.pack_into("<I", data, _fat_offset(data) + 4 * start, start)


def _difat_loop(data: bytearray) -> None:
    struct.pack_into("<II", data, 68, 0, 5)


def _huge_difat_count(data: bytearray) -> None:
    struct.pack_into("<I", data, 72, 0x7FFFFFFF)


def _huge_stream_sizes(data: bytearray) -> None:
    start = struct.unpack_from("<I", data, 48)[0]
    for entry in range(4):
        struct.pack_into("<Q", data, 512 + start * 512 + entry * 128 + 120, 1 << 40)


def _tree_cycles(data: bytearray) -> None:
    start = struct.unpack_from("<I", data, 48)[0]
    for entry in range(4):
        struct.pack_into("<III", data, 512 + start * 512 + entry * 128 + 68, 0, 0, 0)


@pytest.mark.parametrize(
    "damage",
    [
        _fat_self_loops,
        _directory_loop,
        _difat_loop,
        _huge_difat_count,
        _huge_stream_sizes,
        _tree_cycles,
    ],
)
def test_malformed_containers_fail_fast_with_a_clear_error(
    damage: Callable[[bytearray], None],
) -> None:
    data = bytearray(_project("x = 1\n" * 3000).to_bytes())
    damage(data)
    started = time.perf_counter()
    with pytest.raises(WorkbookError):
        ovba.read_project(bytes(data))
    assert time.perf_counter() - started < 2


def test_random_corruption_never_escapes_as_another_error() -> None:
    original = _project("x = 1\n" * 3000).to_bytes()
    generator = random.Random(7)
    started = time.perf_counter()
    for _ in range(400):
        data = bytearray(original)
        for _ in range(generator.randint(1, 4)):
            position = generator.randrange(len(data) if generator.random() < 0.5 else 512)
            data[position : position + 4] = generator.choice(
                [b"\xff\xff\xff\x7f", bytes(4), b"\x01\x00\x00\x00", b"\xfe\xff\xff\xff"]
            )
        with contextlib.suppress(ExcelMCPError):
            ovba.read_project(bytes(data))
    assert time.perf_counter() - started < 30
