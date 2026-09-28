"""Checks on what clients see: schemas, annotations, errors and modes."""

from pathlib import Path
from typing import Any

import pytest
from mcp import Client

from excel_mcp.config import Settings
from excel_mcp.server import create_server
from tests.conftest import error_text

pytestmark = pytest.mark.anyio


def _untyped_schemas(schema: Any, where: str = "") -> list[str]:
    """Find sub-schemas without a type, which clients such as Gemini CLI reject."""
    found: list[str] = []
    if isinstance(schema, dict):
        is_schema = "properties" in schema or "items" in schema or "type" in schema
        has_type = any(key in schema for key in ("type", "anyOf", "oneOf", "allOf", "$ref"))
        if (is_schema and not has_type) or schema == {}:
            found.append(where)
        for key, value in schema.items():
            found += _untyped_schemas(value, f"{where}.{key}")
    elif isinstance(schema, list):
        for index, value in enumerate(schema):
            found += _untyped_schemas(value, f"{where}[{index}]")
    return found


async def test_every_tool_is_documented_and_typed(client: Client) -> None:
    tools = (await client.list_tools()).tools
    assert len(tools) == 26
    for tool in tools:
        assert tool.title, tool.name
        assert tool.description and not tool.description.startswith(" "), tool.name
        assert tool.annotations is not None, tool.name
        assert tool.annotations.open_world_hint is False, tool.name
        assert _untyped_schemas(tool.input_schema) == [], tool.name
        for name, prop in tool.input_schema["properties"].items():
            assert "description" in prop or "$ref" in prop, (tool.name, name)
        for model_name, model in tool.input_schema.get("$defs", {}).items():
            assert "description" in model, (tool.name, model_name)
            for name, prop in model["properties"].items():
                assert "description" in prop, (tool.name, model_name, name)


async def test_tool_order_is_stable(files: Path) -> None:
    names = []
    for _ in range(2):
        async with Client(create_server(Settings(allowed_dirs=[files]))) as client:
            names.append([tool.name for tool in (await client.list_tools()).tools])
    assert names[0] == names[1]


async def test_read_only_mode_registers_only_read_tools(files: Path) -> None:
    async with Client(create_server(Settings(allowed_dirs=[files], read_only=True))) as client:
        tools = (await client.list_tools()).tools
    assert {tool.name for tool in tools} == {
        "describe_workbook",
        "list_workbooks",
        "export_workbook",
        "describe_sheet",
        "read_range",
        "find_cells",
        "read_vba",
    }
    assert all(tool.annotations and tool.annotations.read_only_hint for tool in tools)


async def test_errors_are_tool_errors_with_helpful_text(client: Client) -> None:
    result = await client.call_tool("read_range", {"path": "missing.xlsx", "sheet": "Data"})
    assert result.is_error
    assert "does not exist" in error_text(result)


async def test_invalid_arguments_are_rejected_before_running(client: Client) -> None:
    result = await client.call_tool(
        "insert_rows_or_columns", {"path": "a.xlsx", "sheet": "S", "axis": "diagonal", "at": 1}
    )
    assert result.is_error


async def test_instructions_describe_path_mode(files: Path, tmp_path: Path) -> None:
    confined = create_server(Settings(allowed_dirs=[files]))
    assert "relative" in (confined.instructions or "")
    unconfined = create_server(Settings())
    assert "absolute paths" in (unconfined.instructions or "")
