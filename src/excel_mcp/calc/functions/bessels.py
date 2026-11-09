"""Bessel functions: BESSELJ, BESSELY, BESSELI and BESSELK."""

import math
from collections.abc import Callable

from excel_mcp.calc.bessel import bessel_i, bessel_j, bessel_k, bessel_y
from excel_mcp.calc.registry import function
from excel_mcp.calc.values import NUM, VALUE, FormulaError, Scalar, to_int, to_number


def _bessel(
    name: str, evaluate: Callable[[float, int], float], positive_only: bool, limit: float
) -> None:
    def call(x: Scalar, n: Scalar) -> float:
        if isinstance(x, bool) or isinstance(n, bool):
            raise FormulaError(VALUE)
        argument, order = to_number(x), to_int(n)
        if order < 0 or (positive_only and argument <= 0) or abs(argument) >= limit:
            raise FormulaError(NUM)
        return evaluate(argument, order)

    function(name, kind="scalar")(call)


# Excel refuses arguments from these limits on, for I where its own result would overflow.
_bessel("BESSELJ", bessel_j, positive_only=False, limit=2**27)
_bessel("BESSELI", bessel_i, positive_only=False, limit=709.7000001)
_bessel("BESSELY", bessel_y, positive_only=True, limit=2**27)
_bessel("BESSELK", bessel_k, positive_only=True, limit=math.inf)
