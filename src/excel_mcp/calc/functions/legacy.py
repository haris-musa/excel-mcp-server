"""Compatibility names of functions that Excel 2010 renamed, with the same arguments."""

import statistics

from excel_mcp.calc.functions.distributions import (
    binom_dist,
    confidence_norm,
    expon_dist,
    norm_dist,
    norm_inv,
    norm_s_inv,
    poisson_dist,
)
from excel_mcp.calc.functions.regression import covariance_p
from excel_mcp.calc.registry import function
from excel_mcp.calc.values import Scalar, to_number

function("NORMDIST", kind="scalar")(norm_dist)
function("NORMINV", kind="scalar")(norm_inv)
function("NORMSINV", kind="scalar")(norm_s_inv)
function("POISSON", kind="scalar")(poisson_dist)
function("EXPONDIST", kind="scalar")(expon_dist)
function("BINOMDIST", kind="scalar")(binom_dist)
function("CONFIDENCE", kind="scalar")(confidence_norm)
function("COVAR", array=True)(covariance_p)


@function("NORMSDIST", kind="scalar")
def normsdist(z: Scalar) -> float:
    return statistics.NormalDist().cdf(to_number(z))
