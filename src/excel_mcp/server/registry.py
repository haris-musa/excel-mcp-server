"""Tool registration shared by all tool modules."""

import functools
import inspect
import json
from collections.abc import Callable
from typing import Any, ParamSpec, TypeVar

from mcp.server import MCPServer
from mcp.server.mcpserver.exceptions import ToolError
from mcp.types import CallToolResult, TextContent, ToolAnnotations
from pydantic import BaseModel

from excel_mcp.errors import ExcelMCPError

P = ParamSpec("P")
R = TypeVar("R")

READ_ONLY = ToolAnnotations(
    read_only_hint=True, destructive_hint=False, idempotent_hint=True, open_world_hint=False
)
ADDITIVE = ToolAnnotations(
    read_only_hint=False, destructive_hint=False, idempotent_hint=False, open_world_hint=False
)
DESTRUCTIVE = ToolAnnotations(
    read_only_hint=False, destructive_hint=True, idempotent_hint=False, open_world_hint=False
)


class ToolRegistry:
    """Registers tools on the server; in read-only mode, only read-only tools."""

    def __init__(self, server: MCPServer, read_only: bool) -> None:
        self.server = server
        self.read_only = read_only

    def reader(self, title: str) -> Callable[[Callable[P, R]], Callable[P, R]]:
        return self._register(title, READ_ONLY)

    def writer(self, title: str) -> Callable[[Callable[P, R]], Callable[P, R]]:
        return self._register(title, ADDITIVE)

    def destroyer(self, title: str) -> Callable[[Callable[P, R]], Callable[P, R]]:
        return self._register(title, DESTRUCTIVE)

    def _register(
        self, title: str, annotations: ToolAnnotations
    ) -> Callable[[Callable[P, R]], Callable[P, R]]:
        def decorate(function: Callable[P, R]) -> Callable[P, R]:
            if self.read_only and not annotations.read_only_hint:
                return function
            returns_text = inspect.signature(function).return_annotation is str
            self.server.tool(
                title=title,
                description=inspect.cleandoc(function.__doc__ or ""),
                annotations=annotations,
                structured_output=False if returns_text else None,
            )(_as_tool_errors(function))
            # The SDK's schemas carry generated titles and nullable unions that only cost tokens.
            tool = self.server._tool_manager.get_tool(function.__name__)  # pyright: ignore[reportPrivateUsage]
            assert tool is not None
            _slim_schema(tool.parameters)
            if tool.output_schema:
                _slim_schema(tool.output_schema)
            return function

        return decorate


def _as_tool_errors(function: Callable[P, R]) -> Callable[P, R]:
    """Report expected errors to the model; the SDK hides the details of anything else."""

    @functools.wraps(function)
    def wrapper(*args: P.args, **kwargs: P.kwargs) -> R:
        try:
            return _compact(function(*args, **kwargs))
        except ExcelMCPError as error:
            raise ToolError(str(error)) from error

    return wrapper


def _compact(result: Any) -> Any:
    """Return structured results without default values, and as compact JSON text.

    The SDK's own text rendering is indented JSON that repeats every default.
    """
    if isinstance(result, CallToolResult) or not isinstance(result, BaseModel | dict):
        return result
    data = (
        result
        if isinstance(result, dict)
        else result.model_dump(mode="json", exclude_defaults=True)
    )
    text = json.dumps(data, separators=(",", ":"), ensure_ascii=False)
    return CallToolResult(content=[TextContent(type="text", text=text)], structured_content=data)


def _slim_schema(schema: dict[str, Any]) -> None:
    """Drop generated titles and turn ``X | null`` optional parameters into plain ``X``."""
    schema.pop("title", None)
    for key in ("properties", "$defs"):
        for sub_schema in schema.get(key, {}).values():
            _slim_schema(sub_schema)
    for key in ("items", "additionalProperties"):
        if isinstance(schema.get(key), dict):
            _slim_schema(schema[key])
    for option in schema.get("anyOf", []):
        _slim_schema(option)
    if "anyOf" in schema and schema.get("default", 0) is None:
        remaining = [option for option in schema["anyOf"] if option != {"type": "null"}]
        if len(remaining) == 1:
            del schema["anyOf"], schema["default"]
            schema.update(remaining[0])
