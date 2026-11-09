"""LINEST, LOGEST, TREND and GROWTH: least squares on several variables."""

import math
from dataclasses import dataclass

from excel_mcp.calc.functions.helpers import as_grid
from excel_mcp.calc.leastsquares import Fit, least_squares
from excel_mcp.calc.registry import function
from excel_mcp.calc.values import (
    NA,
    NUM,
    REF,
    VALUE,
    FormulaError,
    Grid,
    Scalar,
    UncalculableError,
    Value,
    is_number,
    to_bool,
)


@dataclass(frozen=True)
class Data:
    y: list[float]
    columns: list[list[float]]  # one list per variable
    by_row: bool  # y is a row, so the variables are rows of known_x

    @property
    def count(self) -> int:
        return len(self.y)


def _numbers(grid: Grid) -> list[list[float]]:
    if not all(is_number(v) for v in grid.flat()):
        raise FormulaError(VALUE)
    return [[float(v) for v in row] for row in grid.rows]  # pyright: ignore[reportArgumentType]


def read_data(known_y: Value, known_x: Value | None) -> Data:
    """The observations, checking shapes before values as Excel does."""
    ys = as_grid(known_y)
    if ys.height != 1 and ys.width != 1:
        raise FormulaError(REF)
    by_row = ys.height == 1 and ys.width > 1
    n = ys.height * ys.width
    xs = None if known_x is None else as_grid(known_x)
    if xs is not None and (xs.width if by_row else xs.height) != n:
        raise FormulaError(REF)
    y = [v for row in _numbers(ys) for v in row]
    if xs is None:
        return Data(y, [[float(i) for i in range(1, n + 1)]], by_row)
    rows = _numbers(xs)
    return Data(y, rows if by_row else [list(c) for c in zip(*rows, strict=True)], by_row)


@dataclass(frozen=True)
class Model:
    fit: Fit
    slopes: list[float]
    intercept: float
    means: list[float]  # mean of each variable, zeros without a constant
    mean_y: float


def _mean(values: list[float]) -> float:
    return math.fsum(values) / len(values)


def _model(data: Data, const: bool) -> Model:
    means = [_mean(c) for c in data.columns] if const else [0.0] * len(data.columns)
    mean_y = _mean(data.y) if const else 0.0
    centered = [[v - m for v in c] for c, m in zip(data.columns, means, strict=True)]
    fit = least_squares(centered, [y - mean_y for y in data.y])
    slopes = fit.coefficients
    intercept = mean_y - math.fsum(s * m for s, m in zip(slopes, means, strict=True))
    return Model(fit, slopes, intercept, means, mean_y)


def _statistics(data: Data, model: Model, const: bool) -> list[list[Scalar]]:
    """The rows below the coefficients, as LINEST lays them out; None marks #N/A."""
    fit = model.fit
    n, variables = data.count, sum(fit.kept)
    if const and any(
        not kept and all(v == c[0] for v in c)
        for kept, c in zip(fit.kept, data.columns, strict=True)
    ):
        raise UncalculableError("LINEST statistics with a constant column")
    df = n - variables - (1 if const else 0)
    ssresid = math.fsum(r * r for r in fit.residual)
    ssreg = math.fsum(
        (y - model.mean_y - r) ** 2 for y, r in zip(data.y, fit.residual, strict=True)
    )
    if df == 0:
        ssresid = 0.0
        ses = [0.0] * len(fit.kept)
        intercept_se = 0.0
        sey, f_value = 0.0, NUM
    else:
        if ssresid == 0:
            raise UncalculableError("LINEST statistics of a perfect fit are rounding noise")
        sigma2 = ssresid / df
        ses = [math.sqrt(sigma2 * fit.inverse_gram[i][i]) for i in range(len(fit.kept))]
        reach = sum(
            a * b * fit.inverse_gram[i][j]
            for i, a in enumerate(model.means)
            for j, b in enumerate(model.means)
        )
        intercept_se = math.sqrt(sigma2 * (1 / n + reach))
        sey = math.sqrt(sigma2)
        f_value = (ssreg / variables) / sigma2 if variables else 0.0
    return [
        [*reversed(ses), intercept_se if const else None],
        [ssreg / (ssreg + ssresid), sey],
        [f_value, float(df)],
        [ssreg, ssresid],
    ]


def _layout(data: Data, const: bool, stats: bool, transform: bool) -> Grid:
    model = _model(data, const)
    coefficients = [*reversed(model.slopes), model.intercept]
    if transform:
        coefficients = [math.exp(v) for v in coefficients]
    rows: list[list[Scalar]] = [[*coefficients]]
    if stats:
        for row in _statistics(data, model, const):
            padded: list[Scalar] = [NA if v is None else v for v in row]
            rows.append(padded + [NA] * (len(coefficients) - len(padded)))
    return Grid(rows)


def _flags(const: Scalar, stats: Scalar) -> tuple[bool, bool]:
    return (True if const is None else to_bool(const)), to_bool(stats)


@function("LINEST", array=True)
def linest(
    known_y: Value, known_x: Value = None, const: Scalar = True, stats: Scalar = False
) -> Grid:
    use_const, show = _flags(const, stats)
    return _layout(read_data(known_y, known_x), use_const, show, transform=False)


def _log_data(data: Data) -> Data:
    if any(y <= 0 for y in data.y):
        raise FormulaError(NUM)
    return Data([math.log(y) for y in data.y], data.columns, data.by_row)


@function("LOGEST", array=True)
def logest(
    known_y: Value, known_x: Value = None, const: Scalar = True, stats: Scalar = False
) -> Grid:
    use_const, show = _flags(const, stats)
    return _layout(_log_data(read_data(known_y, known_x)), use_const, show, transform=True)


def _predict(data: Data, new_x: Value | None, const: bool, grow: bool) -> Grid:
    model = _model(data, const)
    target = data.columns
    shape: list[list[float]] | None = None
    if new_x is not None:
        rows = _numbers(as_grid(new_x))
        if len(data.columns) == 1:
            shape = rows
            target = [[v for row in rows for v in row]]
        else:
            target = rows if data.by_row else [list(c) for c in zip(*rows, strict=True)]
            if len(target) != len(data.columns):
                raise FormulaError(REF)
    values = [
        model.intercept
        + math.fsum(s * column[i] for s, column in zip(model.slopes, target, strict=True))
        for i in range(len(target[0]))
    ]
    if grow:
        values = [math.exp(v) for v in values]
    if shape is not None:
        width = len(shape[0])
        return Grid([values[r * width : (r + 1) * width] for r in range(len(shape))])
    return Grid([values]) if data.by_row else Grid([[v] for v in values])


@function("TREND", array=True)
def trend(known_y: Value, known_x: Value = None, new_x: Value = None, const: Scalar = True) -> Grid:
    use_const, _ = _flags(const, False)
    return _predict(read_data(known_y, known_x), new_x, use_const, grow=False)


@function("GROWTH", array=True)
def growth(
    known_y: Value, known_x: Value = None, new_x: Value = None, const: Scalar = True
) -> Grid:
    use_const, _ = _flags(const, False)
    data = _log_data(read_data(known_y, known_x))
    return _predict(data, new_x, use_const, grow=True)
