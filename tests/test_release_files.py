"""Keep generated docs and release metadata in sync with the code."""

import json
import tomllib
from pathlib import Path
from typing import Any

import generate_tools_doc
import pytest
from mcp.types import Tool

ROOT = Path(__file__).resolve().parent.parent

pytestmark = pytest.mark.anyio


def _version() -> str:
    return tomllib.loads((ROOT / "pyproject.toml").read_text())["project"]["version"]


async def test_tools_doc_is_up_to_date() -> None:
    expected = generate_tools_doc.render(await generate_tools_doc.list_tools())
    actual = (ROOT / "TOOLS.md").read_text(encoding="utf-8")
    assert actual == expected, "Run: uv run python scripts/generate_tools_doc.py"


def _tool(name: str, properties: dict[str, Any], defs: dict[str, Any]) -> Tool:
    schema = {"type": "object", "properties": properties, "required": ["items"], "$defs": defs}
    return Tool(name=name, title=name, description="Summary.", input_schema=schema)


def test_tools_doc_documents_nested_models() -> None:
    value = {"$ref": "#/$defs/Value"}
    defs = {
        "Value": {
            "properties": {
                "field": {"type": "string", "description": "A header."},
                "mode": {"enum": ["sum", "count"], "default": "sum", "type": "string"},
                "group": {"anyOf": [{"$ref": "#/$defs/Group"}, {"type": "null"}]},
            },
            "required": ["field"],
            "type": "object",
        },
        "Group": {"properties": {"by": {"type": "number", "minimum": 1}}, "type": "object"},
    }
    first = _tool("first", {"items": {"type": "array", "items": value}}, defs)
    second = _tool("second", {"item": {"anyOf": [value, {"type": "null"}]}}, defs)
    doc = generate_tools_doc.render([first, second])
    assert "| `items` | array of [Value](#value) | yes |" in doc
    assert "| `item` | [Value](#value) | no |" in doc
    assert "| `field` | string | yes | A header. |" in doc
    assert '| `mode` | `sum` \\| `count` | no | Default: `"sum"`. |' in doc
    assert "| `group` | [Group](#group) | no |" in doc
    assert "| `by` | number | no | Allowed minimum: 1. |" in doc
    assert doc.count("#### Value") == 1
    assert doc.index("## Shared types") < doc.index("#### Value")
    assert doc.index("#### Value") < doc.index("#### Group")


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
