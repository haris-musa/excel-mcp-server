"""Logical functions."""

from typing import TYPE_CHECKING

from excel_mcp.calc.operators import elementwise
from excel_mcp.calc.parser import Name, Node
from excel_mcp.calc.registry import function
from excel_mcp.calc.values import (
    NA,
    VALUE,
    ExcelError,
    FormulaError,
    Grid,
    Scalar,
    UncalculableError,
    Value,
    compare,
    is_number,
    to_bool,
    to_int,
)

if TYPE_CHECKING:
    from excel_mcp.calc.engine import Engine


@function("IF", kind="lazy")
def if_(engine: "Engine", condition: Node, then: Node, otherwise: Node | None = None) -> Value:
    if engine.array_mode:
        return _if_for_arrays(engine, condition, then, otherwise)
    test = engine.scalar(condition)
    if isinstance(test, ExcelError):
        return test
    if to_bool(test):
        return engine.eval(then)
    return engine.eval(otherwise) if otherwise is not None else False


def _if_for_arrays(engine: "Engine", condition: Node, then: Node, otherwise: Node | None) -> Value:
    """IF where a range or array can be the condition: it picks per element."""
    test = engine.eval(condition, True)
    if not isinstance(test, Grid):
        if isinstance(test, ExcelError):
            return test
        branch = then if to_bool(test) else otherwise
        return False if branch is None else engine.eval(branch, True)
    branches = (
        engine.eval(then, True),
        False if otherwise is None else engine.eval(otherwise, True),
    )

    def pick(flag: Scalar, yes: Scalar, no: Scalar) -> Scalar:
        if isinstance(flag, ExcelError):
            return flag
        return yes if to_bool(flag) else no

    return elementwise(pick, test, *branches)


@function("IFS", kind="lazy")
def ifs(engine: "Engine", *pairs: Node) -> Value:
    if len(pairs) % 2:
        raise FormulaError(VALUE)
    for condition, result in zip(pairs[::2], pairs[1::2], strict=True):
        test = engine.scalar(condition)
        if isinstance(test, ExcelError):
            return test
        if to_bool(test):
            return engine.eval(result)
    return NA


@function("IFERROR", kind="lazy")
def iferror(engine: "Engine", value: Node, fallback: Node) -> Value:
    result = engine.scalar(value)
    return engine.eval(fallback) if isinstance(result, ExcelError) else result


@function("IFNA", kind="lazy")
def ifna(engine: "Engine", value: Node, fallback: Node) -> Value:
    result = engine.scalar(value)
    return engine.eval(fallback) if result == NA else result


@function("SWITCH", kind="lazy")
def switch(engine: "Engine", expression: Node, *cases: Node) -> Value:
    target = engine.scalar(expression)
    if isinstance(target, ExcelError):
        return target
    pairs, default = (cases[:-1], cases[-1]) if len(cases) % 2 else (cases, None)
    for candidate, result in zip(pairs[::2], pairs[1::2], strict=True):
        option = engine.scalar(candidate)
        if isinstance(option, ExcelError):
            return option
        if compare(target, option) == 0:
            return engine.eval(result)
    return NA if default is None else engine.eval(default)


@function("CHOOSE", kind="lazy")
def choose(engine: "Engine", index: Node, *options: Node) -> Value:
    position = engine.scalar(index)
    if isinstance(position, ExcelError):
        return position
    number = to_int(position)
    if not 1 <= number <= len(options):
        raise FormulaError(VALUE)
    return engine.eval(options[number - 1])


def _logicals(args: tuple[Value, ...]) -> list[bool]:
    found: list[bool] = []
    for arg in args:
        if isinstance(arg, Grid):
            for item in arg.flat():
                if isinstance(item, ExcelError):
                    raise FormulaError(item)
                if isinstance(item, bool) or is_number(item):
                    found.append(bool(item))
        else:
            found.append(to_bool(arg))
    if not found:
        raise FormulaError(VALUE)
    return found


@function("AND")
def and_(*args: Value) -> bool:
    return all(_logicals(args))


@function("OR")
def or_(*args: Value) -> bool:
    return any(_logicals(args))


@function("XOR")
def xor(*args: Value) -> bool:
    return sum(_logicals(args)) % 2 == 1


@function("NOT", kind="scalar")
def not_(value: Scalar) -> bool:
    return not to_bool(value)


@function("TRUE")
def true() -> bool:
    return True


@function("FALSE")
def false() -> bool:
    return False


@function("LET", kind="lazy")
def let(engine: "Engine", *args: Node) -> Value:
    if len(args) < 3 or len(args) % 2 == 0:
        raise UncalculableError("LET: wrong number of arguments")
    scope: dict[str, Value] = {}
    engine.scopes.append(scope)
    try:
        for name, value in zip(args[:-1:2], args[1:-1:2], strict=True):
            if not isinstance(name, Name):
                raise UncalculableError("LET: not a name")
            scope[name.name.casefold()] = engine.eval(value)
        return engine.eval(args[-1])
    finally:
        engine.scopes.pop()
