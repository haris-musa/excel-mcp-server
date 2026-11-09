"""Number formats for TEXT, DOLLAR and FIXED (US English)."""

from decimal import ROUND_HALF_UP, Decimal

from excel_mcp.calc.numberformat import format_value
from excel_mcp.calc.values import UncalculableError

MAX_FORMAT_LENGTH = 255  # Excel's limit; TEXT gives #VALUE! beyond it


def format_number(value: float, fmt: str) -> str:
    if len(fmt) > MAX_FORMAT_LENGTH:
        raise UncalculableError(f"number format of more than {MAX_FORMAT_LENGTH} characters")
    try:
        return format_value(value, fmt)
    except ArithmeticError:  # more digits than the decimal context or a float can hold
        raise UncalculableError("too many digits for the number format") from None


def format_fixed(value: float, decimals: int, grouping: bool) -> str:
    """Round half away from zero and write with ``decimals`` places."""
    try:
        return _fixed(value, decimals, grouping)
    except ArithmeticError:
        raise UncalculableError("too many digits") from None


def _fixed(value: float, decimals: int, grouping: bool) -> str:
    exact = Decimal(f"{abs(value):.15g}")
    quantum = Decimal(1).scaleb(-decimals)
    rounded = exact.quantize(quantum, rounding=ROUND_HALF_UP) if decimals >= 0 else None
    if rounded is None:
        rounded = (exact.scaleb(decimals).quantize(Decimal(1), rounding=ROUND_HALF_UP)).scaleb(
            -decimals
        )
        decimals = 0
    text = f"{rounded:,.{decimals}f}" if grouping else f"{rounded:.{decimals}f}"
    return ("-" if value < 0 and rounded != 0 else "") + text
