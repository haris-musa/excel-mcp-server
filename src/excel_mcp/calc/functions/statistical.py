"""Statistical functions, including the conditional aggregates (COUNTIF, SUMIFS, ...)."""

import math
import statistics
from typing import TYPE_CHECKING

from excel_mcp.calc.criteria import criterion
from excel_mcp.calc.functions.helpers import array_numbers, as_grid, numbers, pairs
from excel_mcp.calc.functions.math import naive_sum
from excel_mcp.calc.parser import Node, Ref
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
    Scalar,
    UncalculableError,
    Value,
    is_number,
    parse_number,
    scalar,
    to_bool,
    to_int,
    to_number,
)

if TYPE_CHECKING:
    from excel_mcp.calc.engine import Engine


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


def _normal(mean: float, sd: float) -> statistics.NormalDist:
    if sd <= 0:
        raise FormulaError(NUM)
    return statistics.NormalDist(mean, sd)


@function("NORM.S.DIST", kind="scalar")
def norm_s_dist(z: Scalar, cumulative: Scalar) -> float:
    dist = statistics.NormalDist()
    return dist.cdf(to_number(z)) if to_bool(cumulative) else dist.pdf(to_number(z))


@function("NORM.DIST", kind="scalar")
def norm_dist(x: Scalar, mean: Scalar, sd: Scalar, cumulative: Scalar) -> float:
    dist = _normal(to_number(mean), to_number(sd))
    return dist.cdf(to_number(x)) if to_bool(cumulative) else dist.pdf(to_number(x))


@function("NORM.S.INV", kind="scalar")
def norm_s_inv(p: Scalar) -> float:
    return _normal(0, 1).inv_cdf(_probability(p))


@function("NORM.INV", kind="scalar")
def norm_inv(p: Scalar, mean: Scalar, sd: Scalar) -> float:
    return _normal(to_number(mean), to_number(sd)).inv_cdf(_probability(p))


def _probability(p: Scalar) -> float:
    value = to_number(p)
    if not 0 < value < 1:
        raise FormulaError(NUM)
    return value


# -- conditional aggregates ------------------------------------------------------------------


def _reference(value: Value) -> RefGrid:
    if not isinstance(value, RefGrid):
        raise FormulaError(VALUE)
    return value


def _matches(conditions: tuple[Value, ...]) -> list[tuple[int, int]]:
    """Positions where every (range, criterion) pair of ``conditions`` holds."""
    ranges = [_reference(conditions[i]) for i in range(0, len(conditions), 2)]
    tests = [criterion(scalar(conditions[i])) for i in range(1, len(conditions), 2)]
    first = ranges[0]
    if any(r.height != first.height or r.width != first.width for r in ranges):
        raise FormulaError(VALUE)
    return [
        (row, col)
        for row in range(first.height)
        for col in range(first.width)
        if all(test(r.rows[row][col]) for r, test in zip(ranges, tests, strict=True))  # pyright: ignore[reportArgumentType]
    ]


def _selected(grid: Value, positions: list[tuple[int, int]]) -> list[Scalar]:
    cells = _reference(grid)
    return [cells.rows[row][col] for row, col in positions]  # pyright: ignore[reportReturnType]


def _selected_numbers(values: list[Scalar]) -> list[float]:
    for value in values:
        if isinstance(value, ExcelError):
            raise FormulaError(value)
    return [float(v) for v in values if is_number(v)]  # pyright: ignore[reportArgumentType]


@function("COUNTIF")
def countif(cells: Value, crit: Value) -> float:
    return float(len(_matches((cells, crit))))


@function("COUNTIFS")
def countifs(*conditions: Value) -> float:
    if len(conditions) % 2:
        raise UncalculableError("COUNTIFS: wrong number of arguments")
    return float(len(_matches(conditions)))


@function("SUMIFS")
def sumifs(total: Value, *conditions: Value) -> float:
    if len(conditions) % 2 or not conditions:
        raise UncalculableError("SUMIFS: wrong number of arguments")
    values = _selected_numbers(_selected(_same_shape(total, conditions[0]), _matches(conditions)))
    return naive_sum(values)


@function("AVERAGEIFS")
def averageifs(total: Value, *conditions: Value) -> float:
    if len(conditions) % 2 or not conditions:
        raise UncalculableError("AVERAGEIFS: wrong number of arguments")
    values = _selected_numbers(_selected(_same_shape(total, conditions[0]), _matches(conditions)))
    if not values:
        raise FormulaError(DIV0)
    return naive_sum(values) / len(values)


@function("MINIFS")
def minifs(total: Value, *conditions: Value) -> float:
    if len(conditions) % 2 or not conditions:
        raise UncalculableError("MINIFS: wrong number of arguments")
    values = _selected_numbers(_selected(_same_shape(total, conditions[0]), _matches(conditions)))
    return min(values, default=0.0)


@function("MAXIFS")
def maxifs(total: Value, *conditions: Value) -> float:
    if len(conditions) % 2 or not conditions:
        raise UncalculableError("MAXIFS: wrong number of arguments")
    values = _selected_numbers(_selected(_same_shape(total, conditions[0]), _matches(conditions)))
    return max(values, default=0.0)


def _same_shape(total: Value, other: Value) -> Value:
    first, second = _reference(total), _reference(other)
    if first.height != second.height or first.width != second.width:
        raise FormulaError(VALUE)
    return first


def _resized(engine: "Engine", node: Node | None, cells: RefGrid) -> RefGrid:
    """The range ``node`` names, resized to ``cells``' shape from its top-left cell (SUMIF)."""
    if node is None:
        return cells
    if not isinstance(node, Ref) or node.top is None or node.left is None:
        shown = _reference(engine.eval(node))
        if shown.height == cells.height and shown.width == cells.width:
            return shown
        raise UncalculableError("SUMIF with a computed sum range of a different size")
    resized = Ref(
        node.sheets,
        node.top,
        node.left,
        node.top + cells.height - 1,
        node.left + cells.width - 1,
        node.relative,
    )
    return engine.reference(resized)


def _conditional(engine: "Engine", cells: Node, crit: Node, total: Node | None) -> list[Scalar]:
    source = _reference(engine.eval(cells))
    test = criterion(engine.scalar(crit))
    target = _resized(engine, total, source)
    return [
        target.rows[row][col]  # pyright: ignore[reportReturnType]
        for row in range(source.height)
        for col in range(source.width)
        if test(source.rows[row][col])  # pyright: ignore[reportArgumentType]
    ]


@function("SUMIF", kind="lazy")
def sumif(engine: "Engine", cells: Node, crit: Node, total: Node | None = None) -> float:
    return naive_sum(_selected_numbers(_conditional(engine, cells, crit, total)))


@function("AVERAGEIF", kind="lazy")
def averageif(engine: "Engine", cells: Node, crit: Node, total: Node | None = None) -> float:
    values = _selected_numbers(_conditional(engine, cells, crit, total))
    if not values:
        raise FormulaError(DIV0)
    return naive_sum(values) / len(values)


@function("PEARSON", array=True)
def pearson(first: Value, second: Value) -> float:
    return correl(first, second)


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


@function("STANDARDIZE", kind="scalar")
def standardize(x: Scalar, mean: Scalar, sd: Scalar) -> float:
    deviation = to_number(sd)
    if deviation <= 0:
        raise FormulaError(NUM)
    return (to_number(x) - to_number(mean)) / deviation


@function("BINOM.DIST", kind="scalar")
def binom_dist(successes: Scalar, trials: Scalar, p: Scalar, cumulative: Scalar) -> float:
    k, n, probability = to_int(successes), to_int(trials), to_number(p)
    if not (0 <= k <= n and 0 <= probability <= 1):
        raise FormulaError(NUM)

    def mass(i: int) -> float:
        return math.comb(n, i) * probability**i * (1 - probability) ** (n - i)

    return naive_sum(mass(i) for i in range(k + 1)) if to_bool(cumulative) else mass(k)


@function("POISSON.DIST", kind="scalar")
def poisson_dist(x: Scalar, mean: Scalar, cumulative: Scalar) -> float:
    k, rate = to_int(x), to_number(mean)
    if k < 0 or rate < 0:
        raise FormulaError(NUM)

    def mass(i: int) -> float:
        return math.exp(-rate) * rate**i / math.factorial(i)

    return naive_sum(mass(i) for i in range(k + 1)) if to_bool(cumulative) else mass(k)


@function("EXPON.DIST", kind="scalar")
def expon_dist(x: Scalar, rate: Scalar, cumulative: Scalar) -> float:
    value, lam = to_number(x), to_number(rate)
    if value < 0 or lam <= 0:
        raise FormulaError(NUM)
    if to_bool(cumulative):
        return 1 - math.exp(-lam * value)
    return lam * math.exp(-lam * value)


@function("CONFIDENCE.NORM", kind="scalar")
def confidence_norm(alpha: Scalar, sd: Scalar, size: Scalar) -> float:
    a, deviation, n = to_number(alpha), to_number(sd), to_int(size)
    if not 0 < a < 1 or deviation <= 0 or n < 1:
        raise FormulaError(NUM)
    return _normal(0, 1).inv_cdf(1 - a / 2) * deviation / math.sqrt(n)
