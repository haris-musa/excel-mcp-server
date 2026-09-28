"""Tool registration shared by all tool modules."""

import functools
import inspect
from collections.abc import Callable
from typing import ParamSpec, TypeVar

from mcp.server import MCPServer
from mcp.server.mcpserver.exceptions import ToolError
from mcp.types import ToolAnnotations

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
            self.server.tool(
                title=title,
                description=inspect.cleandoc(function.__doc__ or ""),
                annotations=annotations,
            )(_as_tool_errors(function))
            return function

        return decorate


def _as_tool_errors(function: Callable[P, R]) -> Callable[P, R]:
    """Report expected errors to the model; the SDK hides the details of anything else."""

    @functools.wraps(function)
    def wrapper(*args: P.args, **kwargs: P.kwargs) -> R:
        try:
            return function(*args, **kwargs)
        except ExcelMCPError as error:
            raise ToolError(str(error)) from error

    return wrapper
