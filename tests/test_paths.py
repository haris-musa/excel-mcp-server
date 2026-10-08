import os
from pathlib import Path

import pytest
from mcp import Client

from excel_mcp.config import Settings
from excel_mcp.errors import PathNotAllowedError
from excel_mcp.paths import PathPolicy
from excel_mcp.server import create_server
from tests.conftest import error_text

pytestmark = pytest.mark.anyio


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
@pytest.mark.parametrize("raw", ["Q:\\book.xlsx"])
def test_other_drives_and_shares_are_rejected_without_resolving(tmp_path: Path, raw: str) -> None:
    with pytest.raises(PathNotAllowedError, match="outside"):
        PathPolicy([tmp_path]).resolve(raw)


@pytest.mark.skipif(os.name != "nt", reason="drive-relative paths exist on Windows")
@pytest.mark.parametrize("drive", ["Q:", "tmp"])
def test_drive_relative_paths_are_explained(tmp_path: Path, drive: str) -> None:
    drive = tmp_path.drive if drive == "tmp" else drive
    with pytest.raises(PathNotAllowedError, match="relative to a drive"):
        PathPolicy([tmp_path]).resolve(f"{drive}book.xlsx")

NETWORK_AND_DEVICE_PATHS = [
    r"\\10.255.255.1\share\a.xlsx",
    "//10.255.255.1/share/a.xlsx",
    r"\/10.255.255.1/share/a.xlsx",
    r"\\localhost\c$\a.xlsx",
    r"\\?\C:\a.xlsx",
    r"\\?\UNC\10.255.255.1\share\a.xlsx",
    r"\\.\pipe\a.xlsx",
    "//./COM1/a.xlsx",
]


@pytest.mark.parametrize("raw", NETWORK_AND_DEVICE_PATHS)
def test_network_and_device_paths_are_rejected_confined_or_not(tmp_path: Path, raw: str) -> None:
    for policy in (PathPolicy([tmp_path]), PathPolicy([])):
        with pytest.raises(PathNotAllowedError, match=r"(network|device) path"):
            policy.resolve(raw)
        with pytest.raises(PathNotAllowedError, match=r"(network|device) path"):
            policy.resolve_image(raw.replace(".xlsx", ".png"))
        with pytest.raises(PathNotAllowedError, match=r"(network|device) path"):
            policy.resolve_directory(raw)


def test_the_share_of_an_allowed_directory_is_trusted(tmp_path: Path) -> None:
    policy = PathPolicy([tmp_path])
    policy._shares = {r"\\fileserver\team"}  # pyright: ignore[reportPrivateUsage]
    check = policy._check_network_and_device_paths  # pyright: ignore[reportPrivateUsage]
    check(r"\\FILESERVER\Team\reports\a.xlsx")
    check("//fileserver/team/a.xlsx")
    with pytest.raises(PathNotAllowedError, match="network path"):
        check(r"\\other\team\a.xlsx")
    with pytest.raises(PathNotAllowedError, match="device path"):
        check(r"\\?\UNC\fileserver\team\a.xlsx")


@pytest.mark.parametrize(
    "raw",
    [
        "CON",
        "nul.xlsx",
        "Aux.xlsx",
        "prn.txt.xlsx",
        "COM1.xlsx",
        "lpt9.xlsx",
        "com\u00b9.xlsx",
        "reports/NUL.xlsx",
        "con .xlsx",
        "nul..xlsx",
        "CONOUT$.xlsx",
    ],
)
def test_reserved_device_names_are_rejected(tmp_path: Path, raw: str) -> None:
    with pytest.raises(PathNotAllowedError, match="device name"):
        PathPolicy([tmp_path]).resolve(raw)
    with pytest.raises(PathNotAllowedError, match="device name"):
        PathPolicy([]).resolve(str(tmp_path / raw))


@pytest.mark.parametrize("raw", ["console.xlsx", "com10.xlsx", "null.xlsx", "auxiliary/a.xlsx"])
def test_names_that_only_resemble_devices_are_accepted(tmp_path: Path, raw: str) -> None:
    assert PathPolicy([tmp_path]).resolve(raw).name == Path(raw).name


@pytest.mark.parametrize("raw", ["book.xlsx:stream", "book.xlsx::$DATA", "folder:stream/a.xlsx"])
def test_alternate_data_streams_are_rejected_in_every_form(tmp_path: Path, raw: str) -> None:
    for policy in (PathPolicy([tmp_path]), PathPolicy([])):
        with pytest.raises(PathNotAllowedError):
            policy.resolve(raw if policy.confined else str(tmp_path / raw))


async def test_tools_refuse_network_paths_before_touching_them() -> None:
    async with Client(create_server(Settings())) as client:
        result = await client.call_tool(
            "describe_workbook", {"path": r"\\10.255.255.1\share\a.xlsx"}
        )
    assert result.is_error
    assert "network path" in error_text(result)
