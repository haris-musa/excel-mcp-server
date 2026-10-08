"""The table of supported functions and the decorator categories use to join it.

Argument handling is chosen per function:

- ``scalar``: every argument is one value; a range is reduced to the formula cell's
  row or column and, inside an array context, the function is applied element-wise.
- ``check``: like ``scalar`` but error values are passed to the function (ISERROR, N).
- ``eager``: arguments may be ranges or arrays; an error value as an argument is the result.
- ``raw``: like ``eager`` but errors are passed to the function (ISERROR, COUNT).
- ``lazy``: receives the engine and unevaluated argument nodes (IF, IFERROR, LET).

``array`` evaluates all arguments (True) or those at the given positions as arrays, as Excel
does for SUMPRODUCT.
"""

import inspect
from collections.abc import Callable
from dataclasses import dataclass
from typing import Any, Literal

Kind = Literal["scalar", "check", "eager", "raw", "lazy"]


@dataclass(frozen=True)
class Function:
    name: str
    call: Callable[..., Any]
    kind: Kind
    min_args: int
    max_args: float
    array: bool | tuple[int, ...]

    def array_at(self, index: int) -> bool:
        return self.array if isinstance(self.array, bool) else index in self.array


FUNCTIONS: dict[str, Function] = {}


def function(
    *names: str, kind: Kind = "eager", array: bool | tuple[int, ...] = False
) -> Callable[[Callable[..., Any]], Callable[..., Any]]:
    def register(call: Callable[..., Any]) -> Callable[..., Any]:
        parameters = list(inspect.signature(call).parameters.values())
        if kind == "lazy":
            parameters = parameters[1:]
        required = sum(
            p.default is p.empty and p.kind is p.POSITIONAL_OR_KEYWORD for p in parameters
        )
        variadic = any(p.kind is p.VAR_POSITIONAL for p in parameters)
        maximum = float("inf") if variadic else len(parameters)
        for name in names:
            FUNCTIONS[name] = Function(name, call, kind, required, maximum, array)
        return call

    return register
