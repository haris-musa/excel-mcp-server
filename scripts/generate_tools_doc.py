"""Generate TOOLS.md from the tool schemas the server publishes.

Run with ``uv run python scripts/generate_tools_doc.py``; a test fails if TOOLS.md is stale.
"""

import asyncio
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


async def list_tools() -> list[Tool]:
    async with Client(create_server(Settings())) as client:
        return (await client.list_tools()).tools


def render(tools: list[Tool]) -> str:
    lines = [HEADER, "| Tool | Summary |", "| --- | --- |"]
    for tool in tools:
        first_paragraph = (tool.description or "").split("\n\n")[0]
        summary = " ".join(first_paragraph.split())
        lines.append(f"| [`{tool.name}`](#{tool.name}) | {summary} |")
    for tool in tools:
        lines += ["", *_render_tool(tool)]
    return "\n".join(lines) + "\n"


def _render_tool(tool: Tool) -> list[str]:
    hints = tool.annotations
    kind = "read-only" if hints and hints.read_only_hint else "modifies files"
    if hints and hints.destructive_hint:
        kind = "modifies files, may overwrite data"
    lines = [f"## {tool.name}", "", f"**{tool.title}** ({kind})", "", tool.description or ""]
    schema = tool.input_schema
    lines += ["", *_parameter_table(schema, schema.get("$defs", {}))]
    return lines


def _parameter_table(schema: dict[str, Any], defs: dict[str, Any]) -> list[str]:
    required = set(schema.get("required", []))
    rows = ["| Parameter | Type | Required | Description |", "| --- | --- | --- | --- |"]
    nested: list[str] = []
    for name, prop in schema["properties"].items():
        model = _referenced_model(prop, defs)
        description = prop.get("description") or (model or {}).get("description", "")
        if "default" in prop and prop["default"] is not None:
            description += f" Default: `{prop['default']}`."
        rows.append(
            f"| `{name}` | {_type_name(prop, defs)} | {'yes' if name in required else 'no'} "
            f"| {description.strip()} |"
        )
        if model:
            nested += ["", f"`{name}` fields:", "", *_parameter_table(model, defs)]
    return rows + nested


def _referenced_model(prop: dict[str, Any], defs: dict[str, Any]) -> dict[str, Any] | None:
    for option in [prop, *prop.get("anyOf", [])]:
        if "$ref" in option:
            return defs[option["$ref"].rsplit("/", 1)[-1]]
    return None


def _type_name(prop: dict[str, Any], defs: dict[str, Any]) -> str:
    if "$ref" in prop:
        return "object"
    if "enum" in prop:
        return " \\| ".join(f"`{value}`" for value in prop["enum"])
    if "anyOf" in prop:
        options = [_type_name(option, defs) for option in prop["anyOf"]]
        return " \\| ".join(option for option in options if option != "null") or "null"
    if prop.get("type") == "array":
        return f"array of {_type_name(prop['items'], defs)}"
    return prop.get("type", "any")


def main() -> None:
    TOOLS_DOC.write_text(render(asyncio.run(list_tools())), encoding="utf-8", newline="\n")


if __name__ == "__main__":
    main()
