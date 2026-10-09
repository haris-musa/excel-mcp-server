"""The tool surface uses one name per concept, names its units and returns what changed."""

from typing import Any

import pytest
from mcp import Client
from mcp.types import Tool

from excel_mcp.config import Settings
from excel_mcp.server import create_server

pytestmark = pytest.mark.anyio

RETIRED = {
    "start_cell",
    "target_cell",
    "target_sheet",
    "anchor_cell",
    "data_range",
    "source_range",
    "source_sheet",
    "location",
    "data",
    "filters",
    "values",
    "fields",
    "index",
    "table",
}
RETIRED_UNITLESS = {"column_widths", "row_heights", "line_width", "line_weight", "fixed_widths"}
# Notes are changed by a separate pull request.
PROSE_RESULTS = {"set_note", "delete_note"}


async def list_tools() -> list[Tool]:
    server = create_server(Settings(allow_vba_write=True))
    async with Client(server) as client:
        return (await client.list_tools()).tools


def property_names(schema: Any) -> set[str]:
    """Every field name in a JSON schema, however deeply nested."""
    if isinstance(schema, list):
        return set().union(*(property_names(item) for item in schema))
    if not isinstance(schema, dict):
        return set()
    names = set(schema.get("properties", {}))
    return names.union(*(property_names(value) for value in schema.values()))


async def test_retired_parameter_names_are_gone() -> None:
    for tool in await list_tools():
        used = set(tool.input_schema["properties"]) & RETIRED
        used |= property_names(tool.input_schema) & RETIRED_UNITLESS
        assert not used, f"{tool.name} still uses {sorted(used)}"


async def test_sizes_say_their_unit() -> None:
    for tool in await list_tools():
        for name in property_names(tool.input_schema):
            if name.startswith(("width_", "height_")) or name.endswith(
                ("_width", "_height", "_widths", "_heights")
            ):
                assert name.endswith(("_cm", "_pt", "_chars")), f"{tool.name}.{name} has no unit"


async def test_a_cell_is_placed_with_at_and_data_comes_from_source() -> None:
    tools = {tool.name: set(tool.input_schema["properties"]) for tool in await list_tools()}
    placing = ["write_range", "copy_range", "create_chart", "insert_image", "add_slicer"]
    placing += ["create_pivot_table"]
    assert all("at" in tools[name] for name in placing)
    assert all("source" in tools[name] for name in ("create_chart", "create_pivot_table"))
    assert "source" in tools["add_sparklines"] and "range" in tools["add_sparklines"]
    assert {"row_fields", "column_fields", "value_fields", "filter_fields"} <= tools[
        "create_pivot_table"
    ]
    assert "new_name" in tools["create_sheet"] and "new_name" in tools["rename_sheet"]


async def test_every_tool_that_changes_a_workbook_returns_an_object() -> None:
    for tool in await list_tools():
        changes = not (tool.annotations and tool.annotations.read_only_hint)
        if changes and tool.name not in PROSE_RESULTS:
            assert tool.output_schema is not None, f"{tool.name} returns prose"
            assert tool.output_schema["type"] == "object"
