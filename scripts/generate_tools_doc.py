"""Generate TOOLS.md from the tool schemas the server publishes.

Run with ``uv run python scripts/generate_tools_doc.py``; a test fails if TOOLS.md is stale.
"""

import asyncio
import json
from collections import Counter
from pathlib import Path
from typing import Any

from mcp import Client
from mcp.types import Tool

from excel_mcp.config import Settings
from excel_mcp.server import create_server

TOOLS_DOC = Path(__file__).resolve().parent.parent / "TOOLS.md"

HEADER = """# Tools

This file is generated from the server's tool schemas by
`scripts/generate_tools_doc.py`. Do not edit it by hand.

Every tool that takes a `path` accepts `.xlsx`, `.xlsm`, `.xltx` and `.xltm` files.
Cells use A1 notation and row and column numbers are 1-based.
"""

# The shared parameter types carry no description in the schemas, because it would cost
# tokens in every client. The documentation spells them out here instead.
PARAMETER_DESCRIPTIONS = {
    "path": "Workbook file path, relative to the allowed folder or absolute inside it.",
    "sheet": "Sheet name.",
    "range": "A1 range, e.g. `A1:C10`.",
    "at": "Top-left cell, e.g. `B2`.",
    "cell": "A1 cell, e.g. `B2`.",
}


async def list_tools() -> list[Tool]:
    async with Client(create_server(Settings())) as client:
        return (await client.list_tools()).tools


def render(tools: list[Tool]) -> str:
    lines = [HEADER, "| Tool | Summary |", "| --- | --- |"]
    for tool in tools:
        first_paragraph = (tool.description or "").split("\n\n")[0]
        summary = " ".join(first_paragraph.split())
        lines.append(f"| [`{tool.name}`](#{tool.name}) | {summary} |")
    models_by_tool = {tool.name: _models_used(tool.input_schema) for tool in tools}
    users = Counter(name for models in models_by_tool.values() for name in models)
    for tool in tools:
        own = [name for name in models_by_tool[tool.name] if users[name] == 1]
        lines += ["", *_render_tool(tool, own)]
    shared = [name for name, count in users.items() if count > 1]
    if shared:
        defs = {name: model for tool in tools for name, model in _defs(tool).items()}
        lines += ["", "## Shared types", "", "Types used by more than one tool."]
        for name in shared:
            lines += _render_model(name, defs)
    return "\n".join(lines) + "\n"


def _defs(tool: Tool) -> dict[str, Any]:
    return tool.input_schema.get("$defs", {})


def _render_tool(tool: Tool, models: list[str]) -> list[str]:
    hints = tool.annotations
    kind = "read-only" if hints and hints.read_only_hint else "modifies files"
    if hints and hints.destructive_hint:
        kind = "modifies files, may overwrite data"
    lines = [f"## {tool.name}", "", f"**{tool.title}** ({kind})", "", tool.description or ""]
    lines += ["", *_field_table(tool.input_schema, _defs(tool), PARAMETER_DESCRIPTIONS)]
    for name in models:
        lines += _render_model(name, _defs(tool))
    return lines


def _render_model(name: str, defs: dict[str, Any]) -> list[str]:
    lines = ["", f"#### {name}", ""]
    if defs[name].get("description"):
        lines += [defs[name]["description"], ""]
    return lines + _field_table(defs[name], defs, {})


def _field_table(
    schema: dict[str, Any], defs: dict[str, Any], fallback: dict[str, str]
) -> list[str]:
    required = set(schema.get("required", []))
    rows = ["| Field | Type | Required | Description |", "| --- | --- | --- | --- |"]
    for name, prop in schema["properties"].items():
        models = _referenced_models(prop)
        description = prop.get("description") or fallback.get(name, "")
        if not description and len(models) == 1:
            description = defs[models[0]].get("description", "")
        bounds = [f"{key}: {prop[key]}" for key in ("minimum", "maximum") if key in prop]
        if bounds:
            description += f" Allowed {', '.join(bounds)}."
        if prop.get("default") is not None:
            description += f" Default: `{json.dumps(prop['default'])}`."
        rows.append(
            f"| `{name}` | {_type_name(prop)} | {'yes' if name in required else 'no'} "
            f"| {description.strip()} |"
        )
    return rows


def _models_used(schema: dict[str, Any]) -> list[str]:
    """Names of the models a schema reaches, in first-use order, nested ones included."""
    defs = schema.get("$defs", {})
    found: list[str] = []

    def visit(node: dict[str, Any]) -> None:
        for prop in node.get("properties", {}).values():
            for name in _referenced_models(prop):
                if name not in found:
                    found.append(name)
                    visit(defs[name])

    visit(schema)
    return found


def _referenced_models(prop: dict[str, Any]) -> list[str]:
    """Models a property refers to directly, through an array, or through anyOf."""
    if "$ref" in prop:
        return [prop["$ref"].rsplit("/", 1)[-1]]
    nested = [*prop.get("anyOf", []), *([prop["items"]] if "items" in prop else [])]
    return [name for option in nested for name in _referenced_models(option)]


def _type_name(prop: dict[str, Any]) -> str:
    if "$ref" in prop:
        name = prop["$ref"].rsplit("/", 1)[-1]
        return f"[{name}](#{name.lower()})"
    if "enum" in prop:
        return " \\| ".join(f"`{value}`" for value in prop["enum"])
    if "const" in prop:
        return f"`{json.dumps(prop['const'])}`"
    if "anyOf" in prop:
        options = [_type_name(option) for option in prop["anyOf"]]
        return " \\| ".join(option for option in options if option != "null") or "null"
    if prop.get("type") == "array":
        return f"array of {_type_name(prop['items'])}"
    return prop.get("type", "any")


def main() -> None:
    TOOLS_DOC.write_text(render(asyncio.run(list_tools())), encoding="utf-8", newline="\n")


if __name__ == "__main__":
    main()
