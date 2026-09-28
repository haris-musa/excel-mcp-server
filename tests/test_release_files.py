"""Keep generated docs and release metadata in sync with the code."""

import json
import tomllib
from pathlib import Path

import generate_tools_doc
import pytest

ROOT = Path(__file__).resolve().parent.parent

pytestmark = pytest.mark.anyio


def _version() -> str:
    return tomllib.loads((ROOT / "pyproject.toml").read_text())["project"]["version"]


async def test_tools_doc_is_up_to_date() -> None:
    expected = generate_tools_doc.render(await generate_tools_doc.list_tools())
    actual = (ROOT / "TOOLS.md").read_text(encoding="utf-8")
    assert actual == expected, "Run: uv run python scripts/generate_tools_doc.py"


async def test_bundle_manifest_lists_every_tool() -> None:
    manifest = json.loads((ROOT / "manifest.json").read_text())
    tools = await generate_tools_doc.list_tools()
    assert [tool["name"] for tool in manifest["tools"]] == [tool.name for tool in tools]


def test_versions_match() -> None:
    manifest = json.loads((ROOT / "manifest.json").read_text())
    registry = json.loads((ROOT / "server.json").read_text())
    versions = {manifest["version"], registry["version"], registry["packages"][0]["version"]}
    assert versions == {_version()}


def test_changelog_has_an_entry_for_the_version() -> None:
    assert f"## [{_version()}]" in (ROOT / "CHANGELOG.md").read_text(encoding="utf-8")


def test_readme_has_registry_ownership_marker() -> None:
    registry = json.loads((ROOT / "server.json").read_text())
    readme = (ROOT / "README.md").read_text(encoding="utf-8")
    assert f"mcp-name: {registry['name']}" in readme
