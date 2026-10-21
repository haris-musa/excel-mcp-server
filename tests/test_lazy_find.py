# pyright: reportAttributeAccessIssue=false, reportOptionalMemberAccess=false
# The tests look at openpyxl internals and at what the lazy cells hold.
"""Replacing text in a large sheet parses only the rows that may hold it, and gives what
replacing in the whole sheet gives."""

import re
import shutil
from collections.abc import Callable
from pathlib import Path

import pytest

from excel_mcp import lazy_edit
from excel_mcp.cell_budget import CellBudget
from excel_mcp.operations import replace
from tests.conftest import ToolCall, lazily
from tests.package_support import read_parts, sheet_part
from tests.test_lazy_edit import make_book, snapshot
from tests.test_lazy_shared import make_shared_book

pytestmark = pytest.mark.anyio

QUERIES = [
    {"query": "item 3", "replacement": "thing 3"},
    {"query": "ITEM", "replacement": "x"},
    {"query": "ITEM", "replacement": "x", "case_sensitive": True},
    {"query": "item 3", "replacement": "", "exact": True},
    {"query": "text", "replacement": "t"},
    {"query": "2", "replacement": "9"},
    {"query": "B", "replacement": "Q"},
    {"query": "=", "replacement": ""},
    {"query": "1.5", "replacement": "7"},
    {"query": "k7", "replacement": "7"},
    {"query": "A2*2", "replacement": "A2*3", "exact": False},
    {"query": "big", "replacement": "huge"},
    {"query": "big", "replacement": "huge", "in_formulas": False},
    {"query": "$A", "replacement": "$B"},
    {"query": "nothing like this", "replacement": "x"},
]


@pytest.mark.parametrize("arguments", QUERIES, ids=[str(q["query"]) for q in QUERIES])
@pytest.mark.parametrize("make", [make_book, make_shared_book], ids=["plain", "shared"])
async def test_replacing_gives_what_replacing_in_the_whole_sheet_gives(
    call: ToolCall,
    files: Path,
    monkeypatch: pytest.MonkeyPatch,
    make: Callable[[Path], None],
    arguments: dict[str, object],
) -> None:
    make(files / "lazy.xlsx")
    shutil.copy(files / "lazy.xlsx", files / "full.xlsx")

    with lazily(monkeypatch):
        lazy = await call("replace_cells", path="lazy.xlsx", **arguments)
    full = await call("replace_cells", path="full.xlsx", **arguments)

    assert lazy == full
    assert snapshot(files / "lazy.xlsx") == snapshot(files / "full.xlsx")


def test_only_the_blocks_that_may_hold_the_text_are_parsed(
    files: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    make_book(files / "book.xlsx")

    with lazily(monkeypatch):
        book, _ = lazy_edit.load(files / "book.xlsx", CellBudget(3_000_000))
        cells = book["Data"]._cells
        opened = set(cells.loaded)  # the rows of a merge, a link and a note
        replace.replace_cells(
            [book["Data"]],
            "nothing like this",
            "x",
            exact=False,
            case_sensitive=False,
            in_formulas=True,
        )
        assert cells.loaded == opened

        done = replace.replace_cells(
            [book["Data"]], "k7", "z", exact=False, case_sensitive=False, in_formulas=True
        )
        assert done.replaced == {"Data": 1}
        assert cells.loaded == opened | {8 // 4}


def test_text_that_only_the_cells_following_a_master_hold_is_found(
    files: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    """`=B30*C30` is the formula of D30, which the file stores only as 'the same as D2'."""
    make_shared_book(files / "book.xlsx")
    stored = re.search(r"<f t=\"shared\"[^>]*>B2\*C2</f>", _sheet_xml(files / "book.xlsx"))
    assert stored

    with lazily(monkeypatch):
        book, _ = lazy_edit.load(files / "book.xlsx", CellBudget(3_000_000))
        done = replace.replace_cells(
            [book["Data"]], "B30", "B31", exact=False, case_sensitive=False, in_formulas=True
        )
        assert done.replaced == {"Data": 1}
        assert book["Data"]["D30"].value == "=B31*C30"


def _sheet_xml(path: Path) -> str:
    parts = read_parts(path)
    return parts[sheet_part(parts, "Data")].decode()
