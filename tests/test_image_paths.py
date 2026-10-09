"""insert_image reads files from disk, so it must not follow links or aliases out of the root."""

import ctypes
import os
import subprocess
from pathlib import Path

import pytest
from PIL import Image

from tests.conftest import ToolCall

pytestmark = pytest.mark.anyio


@pytest.fixture
def outside(files: Path) -> Path:
    directory = files.parent / "outside"
    directory.mkdir()
    Image.new("RGB", (4, 4), "red").save(directory / "secret.png")
    return directory


async def _insert(call_error: ToolCall, image_path: str) -> str:
    return await call_error(
        "insert_image", path="sales.xlsx", sheet="Report", image_path=image_path, at="A1"
    )


async def test_symlinks_out_of_the_root_are_refused(
    call_error: ToolCall, sample: Path, files: Path, outside: Path
) -> None:
    try:
        (files / "link").symlink_to(outside, target_is_directory=True)
        (files / "file_link.png").symlink_to(outside / "secret.png")
    except OSError:
        pytest.skip("symlinks are not available")
    assert "outside" in await _insert(call_error, "link/secret.png")
    assert "outside" in await _insert(call_error, "file_link.png")


@pytest.mark.skipif(os.name != "nt", reason="junctions are a Windows feature")
async def test_junctions_out_of_the_root_are_refused(
    call_error: ToolCall, sample: Path, files: Path, outside: Path
) -> None:
    result = subprocess.run(
        ["cmd", "/c", "mklink", "/J", str(files / "junction"), str(outside)],
        capture_output=True,
        check=False,
    )
    if result.returncode:
        pytest.skip("junctions could not be created")
    assert "outside" in await _insert(call_error, "junction/secret.png")


@pytest.mark.skipif(os.name != "nt", reason="8.3 short names are a Windows feature")
async def test_short_names_of_outside_folders_are_refused(
    call_error: ToolCall, sample: Path, outside: Path
) -> None:
    buffer = ctypes.create_unicode_buffer(1024)
    ctypes.windll.kernel32.GetShortPathNameW(str(outside), buffer, 1024)  # type: ignore[attr-defined]
    short = buffer.value
    if not short or short.casefold() == str(outside).casefold():
        pytest.skip("8.3 short names are disabled on this volume")
    assert "outside" in await _insert(call_error, f"{short}\\secret.png")


async def test_a_relative_path_cannot_climb_out(
    call_error: ToolCall, sample: Path, outside: Path
) -> None:
    assert "outside" in await _insert(call_error, "../outside/secret.png")
    assert "outside" in await _insert(call_error, str(outside / "secret.png"))
