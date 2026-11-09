"""Statistics on text and logical values, trimmed means, multiple modes and frequencies."""

import math
from bisect import bisect_left

from excel_mcp.calc.functions.helpers import array_numbers, flat, numbers, pairs
from excel_mcp.calc.functions.math import ceiling, naive_sum
from excel_mcp.calc.functions.statistical import a_values
from excel_mcp.calc.registry import function
from excel_mcp.calc.values import (
    DIV0,
    NA,
    NUM,
    ExcelError,
    FormulaError,
    Grid,
    Scalar,
    Value,
    is_number,
    to_number,
)


def _spread(values: list[float], sample: bool) -> float:
    n = len(values)
    if n < (2 if sample else 1):
        raise FormulaError(DIV0)
    mean = naive_sum(values) / n
    return naive_sum((v - mean) ** 2 for v in values) / (n - 1 if sample else n)


@function("VARA")
def vara(*args: Value) -> float:
    return _spread(a_values(args), sample=True)


@function("VARPA")
def varpa(*args: Value) -> float:
    return _spread(a_values(args), sample=False)


@function("STDEVA")
def stdeva(*args: Value) -> float:
    return math.sqrt(_spread(a_values(args), sample=True))


@function("STDEVPA")
def stdevpa(*args: Value) -> float:
    return math.sqrt(_spread(a_values(args), sample=False))


@function("SKEW.P")
def skew_p(*args: Value) -> float:
    values = numbers(args)
    n = len(values)
    if n < 3:
        raise FormulaError(DIV0)
    deviation = math.sqrt(_spread(values, sample=False))
    if deviation == 0:
        raise FormulaError(DIV0)
    mean = naive_sum(values) / n
    return naive_sum(((v - mean) / deviation) ** 3 for v in values) / n


@function("MODE.MULT", array=True)
def mode_mult(*args: Value) -> Grid:
    values = [v for arg in args for v in array_numbers(arg)]
    counts: dict[float, int] = {}
    for value in values:
        counts[value] = counts.get(value, 0) + 1
    best = max(counts.values(), default=0)
    if best < 2:
        raise FormulaError(NA)
    return Grid([[value] for value in counts if counts[value] == best])


@function("TRIMMEAN", array=(0,))
def trimmean(array: Value, percent: Scalar) -> float:
    values = sorted(array_numbers(array))
    p = to_number(percent)
    if not values or not 0 <= p < 1:
        raise FormulaError(NUM)
    trim = int(len(values) * p) // 2
    kept = values[trim : len(values) - trim]
    return naive_sum(kept) / len(kept)


@function("STEYX", array=True)
def steyx(known_y: Value, known_x: Value) -> float:
    ys, xs = pairs(known_y, known_x)
    n = len(xs)
    if n < 3:
        raise FormulaError(DIV0)
    mean_x, mean_y = naive_sum(xs) / n, naive_sum(ys) / n
    sxx = naive_sum((x - mean_x) ** 2 for x in xs)
    syy = naive_sum((y - mean_y) ** 2 for y in ys)
    sxy = naive_sum((x - mean_x) * (y - mean_y) for x, y in zip(xs, ys, strict=True))
    if sxx == 0:
        raise FormulaError(DIV0)
    return math.sqrt(max(0.0, syy - sxy * sxy / sxx) / (n - 2))


@function("FREQUENCY", array=True)
def frequency(data: Value, bins: Value) -> Grid:
    """Counts per bin in the order given: a repeated bin gets 0, and without numeric bins
    Excel counts against one bin of 0."""
    items = flat([data])
    for item in items:
        if isinstance(item, ExcelError):
            raise FormulaError(item)
    values = [float(v) for v in items if is_number(v)]  # pyright: ignore[reportArgumentType]
    edges = [float(v) for v in flat([bins]) if is_number(v)] or [0.0]  # pyright: ignore[reportArgumentType]
    ordered = sorted(set(edges))
    counts = dict.fromkeys(ordered, 0)
    above = 0
    for value in values:
        index = bisect_left(ordered, value)
        if index == len(ordered):
            above += 1
        else:
            counts[ordered[index]] += 1
    seen: set[float] = set()
    rows = []
    for edge in edges:
        rows.append([float(0 if edge in seen else counts[edge])])
        seen.add(edge)
    return Grid([*rows, [float(above)]])


function("ECMA.CEILING", kind="scalar")(ceiling)
