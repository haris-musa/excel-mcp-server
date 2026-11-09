"""The formula text that a sheet's shared formulas expand to when it is loaded.

A shared formula is stored once, as its master, and each other cell of the group names only
the group (`si`). Loading translates the master's text for every one of those cells, so the
work is the length of the master times the cells that use it, however small the file is.
"""

import re
from collections import Counter
from collections.abc import Iterator
from html import unescape
from typing import IO

from excel_mcp.text_budget import TextBudget

_CHUNK = 1 << 20
_MASTER = re.compile(rb'<f\b([^>]*\bt="shared"[^>]*)>([^<]+)</f>')
_USER = re.compile(rb'<f\b[^>]*\bsi="(\d+)"')
_GROUP = re.compile(rb'\bsi="(\d+)"')


def master_lengths(data: bytes) -> dict[bytes, int]:
    """The length of each group's master text; the first master of a group is the one used."""
    lengths: dict[bytes, int] = {}
    for found in _MASTER.finditer(data):
        group = _GROUP.search(found[1])
        if group is not None:
            lengths.setdefault(group[1], len(unescape(found[2].decode())))
    return lengths


def users(data: bytes) -> Counter[bytes]:
    """How many cells of ``data`` name each group."""
    return Counter(_USER.findall(data))


def _pieces(source: IO[bytes]) -> Iterator[bytes]:
    """The XML in parts that end after a cell, so that no formula is cut in two."""
    kept = b""
    while chunk := source.read(_CHUNK):
        kept += chunk
        end = kept.rfind(b"</c>")
        if end >= 0:
            yield kept[: end + 4]
            kept = kept[end + 4 :]
    yield kept


def spend_sheet(source: IO[bytes], budget: TextBudget) -> None:
    """Spend the text that translating the shared formulas of one worksheet's XML produces."""
    lengths: dict[bytes, int] = {}
    counts: Counter[bytes] = Counter()
    for piece in _pieces(source):
        for group, length in master_lengths(piece).items():
            lengths.setdefault(group, length)
        counts.update(users(piece))
    budget.spend(sum(lengths[group] * cells for group, cells in counts.items() if group in lengths))
