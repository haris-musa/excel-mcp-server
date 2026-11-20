"""Work estimates for the functions whose work grows faster than the cells they take and
return; the engine counts those cells itself."""

from collections.abc import Callable

from excel_mcp.calc.values import FormulaError, Grid, UncalculableError, Value, to_number

STEPS_PER_UNIT = 32  # simple steps (a multiply-add, a comparison) counted as one unit of work
_LARGEST = 1e16
_DAYS_PER_YEAR = 365
_COUPON_BASE = 32


def shape(value: Value) -> tuple[int, int]:
    return (value.height, value.width) if isinstance(value, Grid) else (1, 1)


def cells(value: Value) -> int:
    height, width = shape(value)
    return height * width


def magnitude(value: Value) -> float:
    """The size of a number argument for an estimate; 0 for anything that is no number,
    which the function itself then refuses."""
    if isinstance(value, Grid):
        if value.height * value.width != 1:
            return 0.0
        value = value.rows[0][0]
    try:
        return min(abs(to_number(value)), _LARGEST)
    except (FormulaError, UncalculableError, OverflowError):
        return 0.0


def text_length(value: Value) -> int:
    items = value.flat() if isinstance(value, Grid) else [value]
    return sum(len(item) for item in items if isinstance(item, str))


def per_count(index: int, units: float) -> Callable[..., int]:
    """A function that loops as many times as argument ``index`` says (periods, trials)."""

    def cost(*args: Value) -> int:
        return 1 + int(magnitude(args[index]) * units) if index < len(args) else 1

    return cost


def per_day(first: int, second: int, units: float) -> Callable[..., int]:
    """A function that steps over the days between two date arguments."""

    def cost(*args: Value) -> int:
        return 1 + int(abs(magnitude(args[first]) - magnitude(args[second])) * units)

    return cost


def per_coupon(dates: tuple[int, ...], frequency: int, units: int) -> Callable[..., int]:
    """A security function that walks the coupon periods between its dates."""

    def cost(*args: Value) -> int:
        moments = [magnitude(args[i]) for i in dates if i < len(args)]
        per_year = magnitude(args[frequency]) if frequency < len(args) else 1
        periods = (max(moments) - min(moments)) / _DAYS_PER_YEAR * min(max(per_year, 1), 12)
        return _COUPON_BASE + int(periods) * units

    return cost
