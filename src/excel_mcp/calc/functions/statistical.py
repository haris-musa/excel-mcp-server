"""Descriptive statistics, ranking and percentiles."""

import math
import statistics

from excel_mcp.calc.functions.helpers import array_numbers, numbers
from excel_mcp.calc.functions.math import naive_sum
from excel_mcp.calc.registry import function
from excel_mcp.calc.values import (
    DIV0,
    NA,
    NUM,
    VALUE,
    ExcelError,
    FormulaError,
    Grid,
    RefGrid,
    UncalculableError,
    Value,
    is_number,
    parse_number,
    scalar,
    to_int,
    to_number,
)


@function("AVERAGE")
def average(*args: Value) -> float:
    values = numbers(args)
    if not values:
        raise FormulaError(DIV0)
    return naive_sum(values) / len(values)


@function("COUNT", kind="raw")
def count(*args: Value) -> float:
    total = 0
    for arg in args:
        if isinstance(arg, Grid):
            total += sum(is_number(v) for v in arg.flat())
        elif is_number(arg) or isinstance(arg, bool):
            total += 1
        elif isinstance(arg, str):
            total += parse_number(arg) is not None
    return float(total)


@function("COUNTA", kind="raw")
def counta(*args: Value) -> float:
    total = 0
    for arg in args:
        total += sum(v is not None for v in arg.flat()) if isinstance(arg, Grid) else 1
    return float(total)


@function("COUNTBLANK")
def countblank(cells: Value) -> float:
    if not isinstance(cells, RefGrid) or cells.clipped_rows or cells.clipped_cols:
        raise (
            FormulaError(VALUE)
            if not isinstance(cells, RefGrid)
            else UncalculableError("COUNTBLANK of a whole row or column")
        )
    return float(sum(v is None or v == "" for v in cells.flat()))


def _all_values(args: tuple[Value, ...]) -> list[float]:
    """Numbers as AVERAGEA reads them: text in ranges is 0 and TRUE is 1."""
    found: list[float] = []
    for arg in args:
        if isinstance(arg, Grid):
            for v in arg.flat():
                if isinstance(v, ExcelError):
                    raise FormulaError(v)
                if v is not None:
                    found.append(float(v) if isinstance(v, bool | int | float) else 0.0)
        else:
            found.append(to_number(arg))
    return found


@function("AVERAGEA")
def averagea(*args: Value) -> float:
    values = _all_values(args)
    if not values:
        raise FormulaError(DIV0)
    return naive_sum(values) / len(values)


@function("MAXA")
def maxa(*args: Value) -> float:
    return max(_all_values(args), default=0.0)


@function("MINA")
def mina(*args: Value) -> float:
    return min(_all_values(args), default=0.0)


@function("MEDIAN")
def median(*args: Value) -> float:
    values = numbers(args)
    if not values:
        raise FormulaError(NUM)
    return statistics.median(values)


@function("MODE.SNGL", "MODE", array=True)
def mode(*args: Value) -> float:
    values = numbers(args)
    counts: dict[float, int] = {}
    for v in values:
        counts[v] = counts.get(v, 0) + 1
    best = max(counts.values(), default=0)
    if best < 2:
        raise FormulaError(NA)
    return next(v for v in values if counts[v] == best)


def _variance(values: list[float], sample: bool) -> float:
    n = len(values)
    if n < (2 if sample else 1):
        raise FormulaError(DIV0)
    mean = naive_sum(values) / n
    return naive_sum((v - mean) ** 2 for v in values) / (n - 1 if sample else n)


@function("VAR.S", "VAR")
def var_s(*args: Value) -> float:
    return _variance(numbers(args), sample=True)


@function("VAR.P", "VARP")
def var_p(*args: Value) -> float:
    return _variance(numbers(args), sample=False)


@function("STDEV.S", "STDEV")
def stdev_s(*args: Value) -> float:
    return math.sqrt(_variance(numbers(args), sample=True))


@function("STDEV.P", "STDEVP")
def stdev_p(*args: Value) -> float:
    return math.sqrt(_variance(numbers(args), sample=False))


@function("DEVSQ")
def devsq(*args: Value) -> float:
    values = numbers(args)
    if not values:
        raise FormulaError(NUM)
    return _variance(values, sample=False) * len(values)


@function("AVEDEV")
def avedev(*args: Value) -> float:
    values = numbers(args)
    if not values:
        raise FormulaError(NUM)
    mean = naive_sum(values) / len(values)
    return naive_sum(abs(v - mean) for v in values) / len(values)


@function("GEOMEAN")
def geomean(*args: Value) -> float:
    values = numbers(args)
    if not values or any(v <= 0 for v in values):
        raise FormulaError(NUM)
    return math.exp(naive_sum(math.log(v) for v in values) / len(values))


@function("HARMEAN")
def harmean(*args: Value) -> float:
    values = numbers(args)
    if not values or any(v <= 0 for v in values):
        raise FormulaError(NUM)
    return len(values) / naive_sum(1 / v for v in values)


@function("LARGE")
def large(array: Value, k: Value) -> float:
    values = sorted(array_numbers(array), reverse=True)
    position = math.ceil(to_number(scalar(k)))
    if not 1 <= position <= len(values):
        raise FormulaError(NUM)
    return values[position - 1]


@function("SMALL")
def small(array: Value, k: Value) -> float:
    values = sorted(array_numbers(array))
    position = math.ceil(to_number(scalar(k)))
    if not 1 <= position <= len(values):
        raise FormulaError(NUM)
    return values[position - 1]


def _rank(number: Value, ref: Value, order: Value, average: bool) -> float:
    target = to_number(scalar(number))
    if not isinstance(ref, RefGrid):
        raise FormulaError(VALUE)
    values = array_numbers(ref)
    if target not in values:
        raise FormulaError(NA)
    ascending = bool(to_number(scalar(order)))
    ahead = sum((v < target) if ascending else (v > target) for v in values)
    ties = values.count(target)
    return ahead + 1 + ((ties - 1) / 2 if average else 0)


@function("RANK.EQ", "RANK")
def rank_eq(number: Value, ref: Value, order: Value = 0) -> float:
    return _rank(number, ref, order, average=False)


@function("RANK.AVG")
def rank_avg(number: Value, ref: Value, order: Value = 0) -> float:
    return _rank(number, ref, order, average=True)


def _percentile(values: list[float], k: float, exclusive: bool) -> float:
    ordered = sorted(values)
    n = len(ordered)
    if not n:
        raise FormulaError(NUM)
    if exclusive:
        if not 1 / (n + 1) <= k <= n / (n + 1):
            raise FormulaError(NUM)
        position = k * (n + 1) - 1
    else:
        if not 0 <= k <= 1:
            raise FormulaError(NUM)
        position = k * (n - 1)
    low = math.floor(position)
    high = min(low + 1, n - 1)
    return ordered[low] + (position - low) * (ordered[high] - ordered[low])


@function("PERCENTILE.INC", "PERCENTILE")
def percentile_inc(array: Value, k: Value) -> float:
    return _percentile(array_numbers(array), to_number(scalar(k)), exclusive=False)


@function("PERCENTILE.EXC")
def percentile_exc(array: Value, k: Value) -> float:
    return _percentile(array_numbers(array), to_number(scalar(k)), exclusive=True)


@function("QUARTILE.INC", "QUARTILE")
def quartile_inc(array: Value, quart: Value) -> float:
    q = to_int(scalar(quart))
    if not 0 <= q <= 4:
        raise FormulaError(NUM)
    return _percentile(array_numbers(array), q / 4, exclusive=False)


@function("QUARTILE.EXC")
def quartile_exc(array: Value, quart: Value) -> float:
    q = to_int(scalar(quart))
    if not 1 <= q <= 3:
        raise FormulaError(NUM)
    return _percentile(array_numbers(array), q / 4, exclusive=True)


def _percent_rank(array: Value, x: Value, significance: Value, exclusive: bool) -> float:
    values = sorted(array_numbers(array))
    target, digits = to_number(scalar(x)), to_int(scalar(significance))
    n = len(values)
    if digits < 1 or not n:
        raise FormulaError(NUM)
    if target < values[0] or target > values[-1]:
        raise FormulaError(NA)
    below = sum(v < target for v in values)
    if target in values:
        position = float(below)
    else:
        low, high = values[below - 1], values[below]
        position = below - 1 + (target - low) / (high - low)
    if exclusive:
        rank = (position + 1) / (n + 1)
    elif n == 1:
        raise FormulaError(DIV0)
    else:
        rank = position / (n - 1)
    scale = 10**digits
    return math.floor(rank * scale + 1e-9) / scale


@function("PERCENTRANK.INC", "PERCENTRANK")
def percentrank_inc(array: Value, x: Value, significance: Value = 3) -> float:
    return _percent_rank(array, x, significance, exclusive=False)


@function("PERCENTRANK.EXC")
def percentrank_exc(array: Value, x: Value, significance: Value = 3) -> float:
    return _percent_rank(array, x, significance, exclusive=True)


def _standardized(values: list[float], power: int) -> tuple[int, float]:
    n = len(values)
    mean = naive_sum(values) / n
    sd = math.sqrt(naive_sum((v - mean) ** 2 for v in values) / (n - 1))
    if sd == 0:
        raise FormulaError(DIV0)
    return n, naive_sum(((v - mean) / sd) ** power for v in values)


@function("SKEW")
def skew(*args: Value) -> float:
    values = numbers(args)
    if len(values) < 3:
        raise FormulaError(DIV0)
    n, total = _standardized(values, 3)
    return n / ((n - 1) * (n - 2)) * total


@function("KURT")
def kurt(*args: Value) -> float:
    values = numbers(args)
    if len(values) < 4:
        raise FormulaError(DIV0)
    n, total = _standardized(values, 4)
    return n * (n + 1) / ((n - 1) * (n - 2) * (n - 3)) * total - 3 * (n - 1) ** 2 / (
        (n - 2) * (n - 3)
    )
