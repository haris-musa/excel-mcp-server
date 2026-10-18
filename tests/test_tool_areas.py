"""--tools: offering only the core and chosen areas of tools."""

from pathlib import Path

import pytest
from mcp import Client

from excel_mcp import cli
from excel_mcp.config import Settings
from excel_mcp.server import create_server
from excel_mcp.server.areas import AREAS, CORE_TOOLS, parse_areas

pytestmark = pytest.mark.anyio


async def _names(files: Path, **settings: object) -> list[str]:
    server = create_server(Settings(allowed_dirs=[files], **settings))  # type: ignore[arg-type]
    async with Client(server) as client:
        return [tool.name for tool in (await client.list_tools()).tools]


async def test_areas_cover_every_tool_exactly_once(files: Path) -> None:
    """A new tool must be added to an area in server/areas.py."""
    everything = await _names(files, allow_vba_write=True)
    listed = [*CORE_TOOLS, *(name for names in AREAS.values() for name in names)]
    assert sorted(listed) == sorted(everything)


async def test_by_default_every_tool_is_offered(files: Path) -> None:
    assert len(await _names(files)) == len(await _names(files, tool_areas=tuple(AREAS)))


async def test_only_the_core_and_chosen_areas_are_offered_in_the_usual_order(
    files: Path,
) -> None:
    everything = await _names(files)
    names = await _names(files, tool_areas=("charts", "csv"))
    wanted = {*CORE_TOOLS, *AREAS["charts"], *AREAS["csv"]}
    assert names == [name for name in everything if name in wanted]
    assert set(names) == wanted
    assert await _names(files, tool_areas=("charts", "csv")) == names
    assert await _names(files, tool_areas=("csv", "charts")) == names


async def test_areas_combine_with_read_only_mode(files: Path) -> None:
    names = await _names(files, tool_areas=("data", "csv", "pivots"), read_only=True)
    assert set(names) == {*CORE_TOOLS, "find_cells"}


async def test_vba_writing_needs_both_the_flag_and_the_area(files: Path) -> None:
    assert "write_vba_module" not in await _names(files, tool_areas=("vba",))
    assert "write_vba_module" not in await _names(files, tool_areas=("data",), allow_vba_write=True)
    names = await _names(files, tool_areas=("vba",), allow_vba_write=True)
    assert set(names) == {*CORE_TOOLS, *AREAS["vba"]}
    names = await _names(files, tool_areas=("vba",), allow_vba_write=True, read_only=True)
    assert set(names) == {*CORE_TOOLS, "read_vba"}


def test_instructions_mention_vba_only_when_it_is_offered(files: Path) -> None:
    def instructions(**settings: object) -> str:
        server = create_server(Settings(allowed_dirs=[files], **settings))  # type: ignore[arg-type]
        return server.instructions or ""

    assert "VBA writing is enabled" in instructions(allow_vba_write=True)
    assert "VBA writing" not in instructions(allow_vba_write=True, tool_areas=("data",))


async def test_a_restricted_server_works_through_a_client(files: Path, sample: Path) -> None:
    server = create_server(Settings(allowed_dirs=[files], tool_areas=("data",)))
    async with Client(server) as client:
        result = await client.call_tool("describe_workbook", {"path": "sales.xlsx"})
        assert not result.is_error
        result = await client.call_tool("create_chart", {"path": "sales.xlsx"})
        assert result.is_error


def test_area_names_are_parsed_leniently_and_deduplicated() -> None:
    assert parse_areas(" Charts, csv ,charts") == ("charts", "csv")


@pytest.mark.parametrize("text", ["chart", "charts,nope", "", " , "])
def test_unknown_or_missing_areas_list_the_valid_ones(text: str) -> None:
    with pytest.raises(ValueError, match="choose from workbooks, csv, sheets, data"):
        parse_areas(text)


def test_the_flag_and_the_environment_variable_choose_areas(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.delenv("EXCEL_MCP_TOOLS", raising=False)
    assert cli._settings(cli._parser().parse_args(["stdio"])).tool_areas is None
    settings = cli._settings(cli._parser().parse_args(["stdio", "--tools", "charts,csv"]))
    assert settings.tool_areas == ("charts", "csv")

    monkeypatch.setenv("EXCEL_MCP_TOOLS", "pivots")
    assert cli._settings(cli._parser().parse_args(["stdio"])).tool_areas == ("pivots",)
    flag = cli._parser().parse_args(["stdio", "--tools", "names"])
    assert cli._settings(flag).tool_areas == ("names",)

    monkeypatch.setenv("EXCEL_MCP_TOOLS", "")
    assert cli._settings(cli._parser().parse_args(["stdio"])).tool_areas is None


@pytest.mark.parametrize("source", ["flag", "environment"])
def test_unknown_areas_are_a_startup_error_listing_the_valid_ones(
    source: str, monkeypatch: pytest.MonkeyPatch, capsys: pytest.CaptureFixture[str]
) -> None:
    monkeypatch.delenv("EXCEL_MCP_TOOLS", raising=False)
    arguments = ["stdio"]
    if source == "flag":
        arguments += ["--tools", "charts,graphs"]
    else:
        monkeypatch.setenv("EXCEL_MCP_TOOLS", "charts,graphs")
    with pytest.raises(SystemExit):
        cli.main(arguments)
    error = capsys.readouterr().err
    assert "argument --tools: unknown tool area graphs" in error
    assert "choose from workbooks, csv, sheets, data" in error
