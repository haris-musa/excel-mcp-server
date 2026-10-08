"""Information functions: IS*, N, T, TYPE, ERROR.TYPE."""

from excel_mcp.calc.registry import function
from excel_mcp.calc.values import (
    ERROR_NUMBERS,
    NA,
    ExcelError,
    FormulaError,
    Scalar,
    is_number,
    to_number,
)


@function("ISBLANK", kind="check")
def isblank(value: Scalar) -> bool:
    return value is None


@function("ISNUMBER", kind="check")
def isnumber(value: Scalar) -> bool:
    return is_number(value)


@function("ISTEXT", kind="check")
def istext(value: Scalar) -> bool:
    return isinstance(value, str)


@function("ISNONTEXT", kind="check")
def isnontext(value: Scalar) -> bool:
    return not isinstance(value, str)


@function("ISLOGICAL", kind="check")
def islogical(value: Scalar) -> bool:
    return isinstance(value, bool)


@function("ISERROR", kind="check")
def iserror(value: Scalar) -> bool:
    return isinstance(value, ExcelError)


@function("ISERR", kind="check")
def iserr(value: Scalar) -> bool:
    return isinstance(value, ExcelError) and value != NA


@function("ISNA", kind="check")
def isna(value: Scalar) -> bool:
    return value == NA


@function("ISEVEN", kind="scalar")
def iseven(value: Scalar) -> bool:
    return int(to_number(value)) % 2 == 0


@function("ISODD", kind="scalar")
def isodd(value: Scalar) -> bool:
    return int(to_number(value)) % 2 == 1


@function("N", kind="check")
def n(value: Scalar) -> Scalar:
    if isinstance(value, ExcelError):
        return value
    if isinstance(value, bool | int | float):
        return float(value)
    return 0.0


@function("T", kind="check")
def t(value: Scalar) -> Scalar:
    return value if isinstance(value, str | ExcelError) else ""


@function("TYPE", kind="check")
def type_(value: Scalar) -> float:
    if isinstance(value, ExcelError):
        return 16.0
    if isinstance(value, str):
        return 2.0
    return 4.0 if isinstance(value, bool) else 1.0


@function("ERROR.TYPE", kind="check")
def error_type(value: Scalar) -> float:
    if not isinstance(value, ExcelError) or value.code not in ERROR_NUMBERS:
        raise FormulaError(NA)
    return float(ERROR_NUMBERS[value.code])


@function("NA")
def na() -> ExcelError:
    return NA
