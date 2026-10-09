"""Tool registration shared by all tool modules."""

import functools
import inspect
import json
from collections.abc import Callable
from typing import Any, ClassVar, ParamSpec, TypeVar

from mcp.server import MCPServer
from mcp.server.mcpserver.exceptions import ToolError
from mcp.types import CallToolResult, TextContent, ToolAnnotations
from pydantic import BaseModel, ValidationError

from excel_mcp.errors import ExcelMCPError
from excel_mcp.inputs import InputModel
from excel_mcp.server.validation import format_validation_error

_REJECTED = "_rejected_arguments"

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
        return self._register(title, READ_ONLY, advertises_output=True)

    def writer(self, title: str) -> Callable[[Callable[P, R]], Callable[P, R]]:
        return self._register(title, ADDITIVE, advertises_output=False)

    def destroyer(self, title: str) -> Callable[[Callable[P, R]], Callable[P, R]]:
        return self._register(title, DESTRUCTIVE, advertises_output=False)

    def _register(
        self, title: str, annotations: ToolAnnotations, *, advertises_output: bool
    ) -> Callable[[Callable[P, R]], Callable[P, R]]:
        """``advertises_output``: list an output schema. Tools that change a workbook return a
        small receipt that needs none; their results still carry it as structured content."""

        def decorate(function: Callable[P, R]) -> Callable[P, R]:
            if self.read_only and not annotations.read_only_hint:
                return function
            returns_text = inspect.signature(function).return_annotation is str
            returns_text = returns_text or not advertises_output
            self.server.tool(
                title=title,
                description=inspect.cleandoc(function.__doc__ or ""),
                annotations=annotations,
                structured_output=False if returns_text else None,
            )(_as_tool_errors(function))
            # The SDK's schemas carry generated titles and nullable unions that only cost tokens.
            tool = self.server._tool_manager.get_tool(function.__name__)  # pyright: ignore[reportPrivateUsage]
            assert tool is not None
            arguments = tool.fn_metadata.arg_model
            tool.fn_metadata.arg_model = type(
                arguments.__name__,
                (_Validated, InputModel, arguments),
                {"tool_name": function.__name__},
            )
            _slim_schema(tool.parameters)
            if tool.output_schema:
                _slim_schema(tool.output_schema)
            return function

        return decorate


class _Rejected:
    """Stands in for validated arguments; the SDK would print a ValidationError as raw pydantic."""

    def __init__(self, message: str) -> None:
        self.message = message

    def model_dump_one_level(self) -> dict[str, str]:
        return {_REJECTED: self.message}


class _Validated(BaseModel):
    tool_name: ClassVar[str]

    @classmethod
    def model_validate(cls, obj: Any, **kwargs: Any) -> Any:
        try:
            return super().model_validate(obj, **kwargs)
        except ValidationError as error:
            return _Rejected(format_validation_error(cls.tool_name, error))


def _as_tool_errors(function: Callable[P, R]) -> Callable[P, R]:
    """Report expected errors to the model; the SDK hides the details of anything else."""

    @functools.wraps(function)
    def wrapper(*args: P.args, **kwargs: P.kwargs) -> R:
        if _REJECTED in kwargs:
            raise ToolError(kwargs[_REJECTED])
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
    """Drop generated titles and keywords the model does not need, and turn ``X | null`` into ``X``.

    Unknown fields are rejected on the server, an absent optional
    boolean or list is false or empty anyway, and a count or size is obviously not negative.
    """
    schema.pop("title", None)
    if schema.get("additionalProperties") is False:
        del schema["additionalProperties"]
    if schema.get("default") is False or schema.get("default") == []:
        del schema["default"]
    if schema.get("exclusiveMinimum") == 0:
        del schema["exclusiveMinimum"]
    if schema.get("minimum") in (0, 1):
        del schema["minimum"]
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
        del schema["default"]
        if len(remaining) == 1:
            del schema["anyOf"]
            schema.update(remaining[0] | schema)  # the parameter's own description wins
        else:
            schema["anyOf"] = remaining
    _merge_plain_types(schema)


def _merge_plain_types(schema: dict[str, Any]) -> None:
    """Write a union of plain types as one ``type`` list; ``number`` already includes integers."""
    options = schema.get("anyOf", [])
    if not options or any(option.keys() != {"type"} for option in options):
        return
    types = [option["type"] for option in options]
    if "number" in types and "integer" in types:
        types.remove("integer")
    del schema["anyOf"]
    schema["type"] = types[0] if len(types) == 1 else types
