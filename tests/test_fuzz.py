"""Every tool, called with nasty arguments on good and damaged workbooks, ends in a result or a
clear tool error: never "Unexpected error" and never a long wait.

Set FUZZ_ROUNDS (calls per tool and workbook) and FUZZ_SEED to search further than the default.
"""

import gc
import logging
import os
import shutil
import time
import warnings
import zlib
from pathlib import Path

import pytest
from mcp import Client
from mcp.types import TextContent

from excel_mcp.config import Settings
from excel_mcp.server import create_server
from tests.fuzz_support import (
    DAMAGE,
    FIXTURES,
    SEEDS,
    Arguments,
    damaged_copies,
    large_merge_workbook,
    rich_workbook,
    sample_files,
    sheet_names,
)

pytestmark = pytest.mark.anyio

ROUNDS = int(os.environ.get("FUZZ_ROUNDS", "1"))
SEED = int(os.environ.get("FUZZ_SEED", "1"))
SLOW_SECONDS = 5
# Workbooks that tools open for real; each is tried alone, damaged copies of the first few too.
DAMAGED = ["rich.xlsx", "macro01.xlsm"]


LABELS = [
    "rich.xlsx",
    "big_merge.xlsx",
    *SEEDS,
    *(f"{Path(name).stem}_{kind}.xlsx" for name in DAMAGED for kind in DAMAGE),
]


@pytest.fixture(scope="module")
def seeds(tmp_path_factory: pytest.TempPathFactory) -> dict[str, Path]:
    directory = tmp_path_factory.mktemp("seeds")
    rich_workbook(directory / "rich.xlsx")
    large_merge_workbook(directory / "big_merge.xlsx")
    found = {"rich.xlsx": directory / "rich.xlsx", "big_merge.xlsx": directory / "big_merge.xlsx"}
    for name in SEEDS:
        found[name] = Path(shutil.copy(FIXTURES / name, directory / name))
    for name in DAMAGED:
        for copy in damaged_copies(found[name], directory).values():
            found[copy.name] = copy
    return found


def _exception(record: logging.LogRecord) -> str:
    assert record.exc_info is not None
    frame = record.exc_info[2]
    while frame is not None and frame.tb_next is not None:
        frame = frame.tb_next
    where = f" at {Path(frame.tb_frame.f_code.co_filename).name}:{frame.tb_lineno}" if frame else ""
    kind, error = record.exc_info[0], record.exc_info[1]
    return f" {kind.__name__ if kind else ''}: {str(error)[:100]}{where}"


@pytest.mark.parametrize("label", LABELS)
async def test_every_tool_survives_nasty_arguments(
    label: str, seeds: dict[str, Path], tmp_path: Path, caplog: pytest.LogCaptureFixture
) -> None:
    workbook = seeds[label]
    work = tmp_path / "work"
    work.mkdir()
    files = sample_files(work)
    name = "book" + workbook.suffix
    sheets = sheet_names(workbook)
    server = create_server(Settings(allowed_dirs=[work], allow_vba_write=True))
    failures: list[str] = []
    async with Client(server) as client:
        for tool in (await client.list_tools()).tools:
            for attempt in range(ROUNDS):
                seed = zlib.crc32(f"{SEED}/{label}/{tool.name}/{attempt}".encode())
                # A folder for each call: Windows refuses to delete a file it is still scanning.
                folder = work / f"call{len(failures)}_{tool.name}_{attempt}"
                folder.mkdir()
                shutil.copy(workbook, folder / name)
                arguments = Arguments(seed, f"{folder.name}/{name}", files, sheets).make(
                    tool.input_schema
                )
                caplog.clear()
                started = time.perf_counter()
                with warnings.catch_warnings(record=True) as caught:
                    warnings.simplefilter("always")
                    result = await client.call_tool(tool.name, arguments)
                    if result.is_error:
                        gc.collect()  # a file left open is reported when it is collected
                elapsed = time.perf_counter() - started
                text = "".join(c.text for c in result.content if isinstance(c, TextContent))
                leaked = [str(w.message)[:90] for w in caught if w.category is ResourceWarning]
                if "Unexpected error" in text or elapsed > SLOW_SECONDS or leaked:
                    logged = [r.getMessage() + _exception(r) for r in caplog.records if r.exc_info]
                    failures.append(
                        f"{label} {tool.name} {arguments!r:.300} -> {text[:60]!r} {elapsed:.1f}s "
                        f"{logged[:1]} {leaked[:1]}"
                    )
    assert not failures, "\n".join(failures)
