"""Correlation, covariance and linear regression."""

import math

from excel_mcp.calc.functions.helpers import as_grid, pairs
from excel_mcp.calc.functions.math import naive_sum
from excel_mcp.calc.registry import function
from excel_mcp.calc.values import (
    DIV0,
    FormulaError,
    Grid,
    UncalculableError,
    Value,
    scalar,
    to_number,
)


def _moments(ys: Value, xs: Value) -> tuple[list[float], list[float], float, float]:
    y, x = pairs(ys, xs)
    if len(x) < 2:
        raise FormulaError(DIV0)
    mean_x, mean_y = naive_sum(x) / len(x), naive_sum(y) / len(y)
    return x, y, mean_x, mean_y


def _sums(ys: Value, xs: Value) -> tuple[float, float, float, float, float]:
    x, y, mean_x, mean_y = _moments(ys, xs)
    sxx = naive_sum((a - mean_x) ** 2 for a in x)
    syy = naive_sum((b - mean_y) ** 2 for b in y)
    sxy = naive_sum((a - mean_x) * (b - mean_y) for a, b in zip(x, y, strict=True))
    return sxx, syy, sxy, mean_x, mean_y


@function("CORREL", array=True)
def correl(first: Value, second: Value) -> float:
    sxx, syy, sxy, _, _ = _sums(first, second)
    if sxx == 0 or syy == 0:
        raise FormulaError(DIV0)
    return sxy / math.sqrt(sxx * syy)


@function("RSQ", array=True)
def rsq(known_y: Value, known_x: Value) -> float:
    sxx, syy, sxy, _, _ = _sums(known_y, known_x)
    if sxx == 0 or syy == 0:
        raise FormulaError(DIV0)
    return sxy * sxy / (sxx * syy)


@function("COVARIANCE.S", array=True)
def covariance_s(first: Value, second: Value) -> float:
    _, _, sxy, _, _ = _sums(first, second)
    return sxy / (len(pairs(first, second)[0]) - 1)


@function("COVARIANCE.P", array=True)
def covariance_p(first: Value, second: Value) -> float:
    x, y = pairs(first, second)
    if not x:
        raise FormulaError(DIV0)
    mean_x, mean_y = naive_sum(x) / len(x), naive_sum(y) / len(y)
    return naive_sum((a - mean_x) * (b - mean_y) for a, b in zip(x, y, strict=True)) / len(x)


@function("SLOPE", array=True)
def slope(known_y: Value, known_x: Value) -> float:
    sxx, _, sxy, _, _ = _sums(known_y, known_x)
    if sxx == 0:
        raise FormulaError(DIV0)
    return sxy / sxx


@function("INTERCEPT", array=True)
def intercept(known_y: Value, known_x: Value) -> float:
    sxx, _, sxy, mean_x, mean_y = _sums(known_y, known_x)
    if sxx == 0:
        raise FormulaError(DIV0)
    return mean_y - sxy / sxx * mean_x


@function("FORECAST.LINEAR", "FORECAST", array=(1, 2))
def forecast(x: Value, known_y: Value, known_x: Value) -> float:
    sxx, _, sxy, mean_x, mean_y = _sums(known_y, known_x)
    if sxx == 0:
        raise FormulaError(DIV0)
    return mean_y + sxy / sxx * (to_number(scalar(x)) - mean_x)


@function("TREND", array=True)
def trend(known_y: Value, known_x: Value = None, new_x: Value = None) -> Grid:
    ys = as_grid(known_y)
    if known_x is None:
        known_x = Grid([[float(i)] for i in range(1, ys.height * ys.width + 1)])
    if as_grid(known_x).width != 1 or ys.width != 1:
        raise UncalculableError("TREND with several columns")
    sxx, _, sxy, mean_x, mean_y = _sums(known_y, known_x)
    if sxx == 0:
        raise FormulaError(DIV0)
    points = as_grid(known_x if new_x is None else new_x)
    rows = [[mean_y + sxy / sxx * (to_number(v) - mean_x) for v in row] for row in points.rows]
    return Grid(rows)  # pyright: ignore[reportArgumentType]


@function("PEARSON", array=True)
def pearson(first: Value, second: Value) -> float:
    return correl(first, second)
