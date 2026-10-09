import os
from pathlib import Path

import pytest

from excel_mcp.errors import PathNotAllowedError
from excel_mcp.paths import PathPolicy


def test_relative_paths_resolve_in_first_allowed_dir(tmp_path: Path) -> None:
    policy = PathPolicy([tmp_path])
    assert policy.resolve("reports/q1.xlsx") == tmp_path.resolve() / "reports" / "q1.xlsx"


def test_absolute_path_inside_allowed_dir_is_accepted(tmp_path: Path) -> None:
    policy = PathPolicy([tmp_path])
    assert policy.resolve(str(tmp_path / "a.xlsx")) == tmp_path.resolve() / "a.xlsx"


@pytest.mark.parametrize("raw", ["../outside.xlsx", "a/../../outside.xlsx"])
def test_traversal_is_rejected(tmp_path: Path, raw: str) -> None:
    with pytest.raises(PathNotAllowedError, match="outside"):
        PathPolicy([tmp_path / "root"]).resolve(raw)


def test_absolute_path_outside_allowed_dir_is_rejected(tmp_path: Path) -> None:
    with pytest.raises(PathNotAllowedError, match="outside"):
        PathPolicy([tmp_path / "root"]).resolve(str(tmp_path / "other.xlsx"))


def test_symlink_escape_is_rejected(tmp_path: Path) -> None:
    root = tmp_path / "root"
    root.mkdir()
    outside = tmp_path / "outside"
    outside.mkdir()
    try:
        (root / "link").symlink_to(outside, target_is_directory=True)
    except OSError:
        pytest.skip("symlinks are not available")
    with pytest.raises(PathNotAllowedError, match="outside"):
        PathPolicy([root]).resolve("link/book.xlsx")


@pytest.mark.parametrize("raw", ["notes.txt", "script.py", "book.xls", "book.csv", "book"])
def test_non_excel_files_are_rejected(tmp_path: Path, raw: str) -> None:
    with pytest.raises(PathNotAllowedError, match="Excel file"):
        PathPolicy([tmp_path]).resolve(raw)


@pytest.mark.parametrize("raw", ["", "   ", "a\x00b.xlsx"])
def test_empty_and_nul_paths_are_rejected(tmp_path: Path, raw: str) -> None:
    with pytest.raises(PathNotAllowedError):
        PathPolicy([tmp_path]).resolve(raw)


def test_unconfined_policy_requires_absolute_paths(tmp_path: Path) -> None:
    policy = PathPolicy([])
    assert policy.resolve(str(tmp_path / "a.xlsx")) == (tmp_path / "a.xlsx").resolve()
    with pytest.raises(PathNotAllowedError, match="absolute"):
        policy.resolve("a.xlsx")


def test_display_is_relative_to_first_allowed_dir(tmp_path: Path) -> None:
    second = tmp_path / "second"
    policy = PathPolicy([tmp_path / "first", second])
    assert policy.display(policy.resolve("sub/a.xlsx")) == "sub/a.xlsx"
    assert policy.display(policy.resolve(str(second / "b.xlsx"))) == str(
        second.resolve() / "b.xlsx"
    )


def test_resolve_directory_defaults_to_first_allowed_dir(tmp_path: Path) -> None:
    assert PathPolicy([tmp_path]).resolve_directory("") == tmp_path.resolve()


def test_resolve_directory_is_confined(tmp_path: Path) -> None:
    with pytest.raises(PathNotAllowedError, match="outside"):
        PathPolicy([tmp_path / "root"]).resolve_directory(os.fspath(tmp_path))


def test_alternate_data_streams_are_rejected(tmp_path: Path) -> None:
    with pytest.raises(PathNotAllowedError, match="':'"):
        PathPolicy([tmp_path]).resolve("notes.txt:hidden.xlsx")


@pytest.mark.skipif(os.name != "nt", reason="drives and UNC shares are Windows paths")
@pytest.mark.parametrize("raw", ["\\\\10.255.255.1\\share\\a.xlsx", "Q:\\book.xlsx"])
def test_other_drives_and_shares_are_rejected_without_resolving(tmp_path: Path, raw: str) -> None:
    with pytest.raises(PathNotAllowedError, match="outside"):
        PathPolicy([tmp_path]).resolve(raw)


@pytest.mark.skipif(os.name != "nt", reason="drive-relative paths exist on Windows")
@pytest.mark.parametrize("drive", ["Q:", "tmp"])
def test_drive_relative_paths_are_explained(tmp_path: Path, drive: str) -> None:
    drive = tmp_path.drive if drive == "tmp" else drive
    with pytest.raises(PathNotAllowedError, match="relative to a drive"):
        PathPolicy([tmp_path]).resolve(f"{drive}book.xlsx")
