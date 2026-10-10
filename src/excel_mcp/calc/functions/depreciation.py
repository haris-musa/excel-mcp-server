"""Depreciation functions: SLN, SYD, DB, DDB, VDB."""

from excel_mcp.calc.registry import function
from excel_mcp.calc.values import (
    DIV0,
    NUM,
    FormulaError,
    Scalar,
    UncalculableError,
    check_periods,
    to_number,
)


@function("SLN", kind="scalar")
def sln(cost: Scalar, salvage: Scalar, life: Scalar) -> float:
    n = to_number(life)
    if n == 0:
        raise FormulaError(DIV0)
    return (to_number(cost) - to_number(salvage)) / n


@function("SYD", kind="scalar")
def syd(cost: Scalar, salvage: Scalar, life: Scalar, per: Scalar) -> float:
    c, s, n, p = to_number(cost), to_number(salvage), to_number(life), to_number(per)
    if n <= 0 or p <= 0 or p > n or s < 0:
        raise FormulaError(NUM)
    return (c - s) * (n - p + 1) * 2 / (n * (n + 1))


@function("DB", kind="scalar")
def db(cost: Scalar, salvage: Scalar, life: Scalar, period: Scalar, month: Scalar = 12) -> float:
    c, s, n, p, m = (to_number(v) for v in (cost, salvage, life, period, month))
    if c < 0 or s < 0 or n <= 0 or p <= 0 or not 1 <= m <= 12 or p > n + (m < 12):
        raise FormulaError(NUM)
    if c == 0:
        return 0.0
    if p != int(p) or n != int(n) or m != int(m):
        raise UncalculableError("DB with fractional arguments")
    check_periods(p)
    rate = round(1 - (s / c) ** (1 / n), 3)
    total = 0.0
    depreciation = 0.0
    for period_ in range(1, int(p) + 1):
        if period_ == 1:
            depreciation = c * rate * m / 12
        elif period_ == n + 1:
            depreciation = (c - total) * rate * (12 - m) / 12
        else:
            depreciation = (c - total) * rate
        total += depreciation
    return depreciation


def _declining(
    cost: float, salvage: float, life: float, start: float, end: float, factor: float, switch: bool
) -> float:
    if start != int(start) or end != int(end) or life != int(life):
        raise UncalculableError("depreciation over fractional periods")
    check_periods(end)
    total = taken = 0.0
    for period in range(1, int(end) + 1):
        declining = min((cost - total) * factor / life, max(cost - salvage - total, 0.0))
        straight = (cost - total - salvage) / (life - period + 1) if switch else 0.0
        amount = max(declining, straight) if switch else declining
        amount = max(min(amount, cost - salvage - total), 0.0)
        if period > start:
            taken += amount
        total += amount
    return taken


@function("DDB", kind="scalar")
def ddb(cost: Scalar, salvage: Scalar, life: Scalar, period: Scalar, factor: Scalar = 2) -> float:
    c, s, n, p, f = (to_number(v) for v in (cost, salvage, life, period, factor))
    if c < 0 or s < 0 or n <= 0 or p <= 0 or p > n or f <= 0:
        raise FormulaError(NUM)
    return _declining(c, s, n, p - 1, p, f, switch=False)


@function("VDB", kind="scalar")
def vdb(
    cost: Scalar,
    salvage: Scalar,
    life: Scalar,
    start: Scalar,
    end: Scalar,
    factor: Scalar = 2,
    no_switch: Scalar = False,
) -> float:
    c, s, n, a, b, f = (to_number(v) for v in (cost, salvage, life, start, end, factor))
    if c < 0 or s < 0 or n <= 0 or a < 0 or b < a or b > n or f <= 0:
        raise FormulaError(NUM)
    return _declining(c, s, n, a, b, f, switch=not to_number(no_switch))
