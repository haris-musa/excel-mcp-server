"""The calculator's distributions, regression, database, engineering and security functions.

Excel's own results for each of these are pinned by formula_golden.json; these tests state the
values a reader can check by hand, and the places where the calculator declines to answer.
"""

import json
from pathlib import Path

import pytest

from excel_mcp.calc.registry import FUNCTIONS
from excel_mcp.calc.values import UncalculableError
from tests.test_calc_parity import calculate


def test_argument_counts_cover_what_excel_accepts() -> None:
    known = json.loads(
        (Path(__file__).parent.parent / "src" / "excel_mcp" / "function_arguments.json").read_text()
    )
    narrower = {
        name: (function.min_args, function.max_args, known[name])
        for name, function in FUNCTIONS.items()
        if name in known
        and (function.min_args > known[name][0] or function.max_args < min(known[name][1], 254))
    }
    # INDEX's fourth argument selects an area of a union; Excel parses BYROW(x) and BYCOL(x).
    assert set(narrower) <= {"INDEX", "BYROW", "BYCOL"}


@pytest.mark.parametrize(
    ("formula", "expected"),
    [
        ("=CHISQ.INV.RT(0.05,10)", 18.307038053275146),
        ("=CHISQ.DIST(3.5,4,TRUE)", 0.522121655511276),
        ("=T.INV.2T(0.05,10)", 2.2281388519649385),
        ("=T.DIST.RT(2,10)", 0.036694018),
        ("=F.INV.RT(0.05,3,5)", 5.409451318117446),
        ("=GAMMA.INV(0.5,3,2)", 5.348120627447115),
        ("=BETA.DIST(0.5,2,3,TRUE)", 0.6875),
        ("=BETA.INV(0.6875,2,3)", 0.5),
        ("=WEIBULL.DIST(2,2,3,TRUE)", 0.3588196115700454),
        ("=HYPGEOM.DIST(1,4,8,20,FALSE)", 0.3632610939112487),
        ("=NEGBINOM.DIST(3,2,0.5,FALSE)", 0.125),
        ("=BINOM.INV(10,0.5,0.5)", 5),
        ("=BINOM.DIST.RANGE(10,0.5,3,5)", 0.568359375),
        ("=CONFIDENCE.T(0.05,2,10)", 1.430713811941329),
        ("=Z.TEST({1,2,3,4,5},3)", 0.5),
    ],
)
def test_distributions(formula: str, expected: float) -> None:
    assert calculate(formula) == pytest.approx(expected, rel=1e-6)


def test_excels_distribution_limits() -> None:
    assert calculate("=CHISQ.DIST(3,10000000001,TRUE)") == "#NUM!"
    assert calculate("=T.DIST(1,0,FALSE)") == "#DIV/0!"
    assert calculate("=BETA.DIST(0,1,1,FALSE)") == "#NUM!"
    assert calculate("=BETA.INV(1,2,3)") == "#NUM!"
    assert calculate("=NEGBINOM.DIST(1,1,1,TRUE)") == "#NUM!"
    assert calculate("=PROB({1,2},{0.5,0.5},1)") == 0.5
    assert calculate("=PROB({1,2},{0.5,0.4},1)") == "#NUM!"


def test_an_f_distribution_with_huge_degrees_of_freedom_keeps_its_digits() -> None:
    assert calculate("=F.DIST.RT(2.5,5,10000000000)") == pytest.approx(
        0.028543123393529754, rel=1e-9
    )
    assert calculate("=T.DIST.RT(1.5,10000000000)") == pytest.approx(0.06680720128464304, rel=1e-9)


def test_chi_square_test_skips_pairs_with_a_blank() -> None:
    cells = {"A1": 10, "B1": 20, "C1": 30, "A2": 12, "B2": 18, "C2": 28}
    assert calculate("=CHISQ.TEST(A1:C1,A2:C2)", cells) == pytest.approx(
        calculate("=CHISQ.DIST.RT(4/12+4/18+4/28,2)"), rel=1e-12
    )


def test_linest_returns_the_coefficients_in_reverse_with_statistics() -> None:
    cells = {f"A{i}": 2 * i + 1 + (i % 2) / 10 for i in range(1, 6)}
    cells |= {f"B{i}": i for i in range(1, 6)}
    assert calculate("=INDEX(LINEST(A1:A5,B1:B5),1,1)", cells) == pytest.approx(2, abs=0.1)
    assert calculate("=INDEX(LINEST(A1:A5,B1:B5,TRUE,TRUE),4,2)", cells) == 3
    assert calculate("=INDEX(LINEST(A1:A5,B1:B5,FALSE,TRUE),1,2)", cells) == 0
    assert 0.99 < calculate("=INDEX(LINEST(A1:A5,B1:B5,TRUE,TRUE),3,1)", cells) <= 1


def test_the_statistics_of_a_perfect_fit_are_left_uncalculated() -> None:
    cells = {f"A{i}": 2 * i + 1 for i in range(1, 6)} | {f"B{i}": i for i in range(1, 6)}
    assert calculate("=INDEX(LINEST(A1:A5,B1:B5),1,1)", cells) == pytest.approx(2)
    assert calculate("=INDEX(LINEST(A1:A5,B1:B5),1,2)", cells) == pytest.approx(1)
    with pytest.raises(UncalculableError):
        calculate("=INDEX(LINEST(A1:A5,B1:B5,TRUE,TRUE),4,1)", cells)


def test_linest_drops_a_repeated_column_and_keeps_the_larger_one() -> None:
    cells = {f"A{i}": y for i, y in enumerate([1.5, 3.1, 4.4, 6.2, 7.1], 1)}
    cells |= {f"B{i}": i for i in range(1, 6)} | {f"C{i}": 2 * i for i in range(1, 6)}
    assert calculate("=INDEX(LINEST(A1:A5,B1:C5),1,1)", cells) != 0
    assert calculate("=INDEX(LINEST(A1:A5,B1:C5),1,2)", cells) == 0
    assert calculate("=INDEX(LINEST(A1:A5,B1:C5,TRUE,TRUE),4,2)", cells) == 3


def test_regression_shapes_and_errors() -> None:
    cells = {f"A{i}": i * i for i in range(1, 5)} | {f"B{i}": i for i in range(1, 5)}
    assert calculate("=ROWS(LINEST(A1:A4,B1:B4,TRUE,TRUE))", cells) == 5
    assert calculate("=COLUMNS(LINEST(A1:A4,B1:B4,TRUE,TRUE))", cells) == 2
    assert calculate("=INDEX(LINEST(A1:A4,B1:B3),1,1)", cells) == "#REF!"
    assert calculate("=INDEX(LOGEST({1,2,0,4},{1,2,3,4}),1,1)") == "#NUM!"
    assert calculate("=INDEX(GROWTH(A1:A4,B1:B4,{5}),1,1)", cells) == pytest.approx(
        calculate("=EXP(INDEX(TREND(LN(A1:A4),B1:B4,{5}),1,1))", cells)
    )


def test_statistics_on_text_and_logical_values() -> None:
    cells = {"A1": 1, "A2": "x", "A3": True, "A4": 4}
    assert calculate("=AVERAGEA(A1:A4)", cells) == 1.5
    assert calculate("=VARA(A1:A4)", cells) == pytest.approx(3.0)
    assert calculate("=STDEVPA(A1:A4)", cells) == pytest.approx(1.5)
    assert calculate("=AVERAGEA({1,TRUE,4})") == 2.5
    assert calculate("=SKEW.P(1,2,9)") == pytest.approx(0.6654688661238354)
    assert calculate("=TRIMMEAN({1,2,3,4,5,6,7,8,9,10},0.2)") == 5.5
    assert calculate("=INDEX(FREQUENCY({1,2,3,4},{2,1}),1)") == 1
    assert calculate("=ROWS(MODE.MULT({1,1,2,2}))") == 2


def test_database_functions_use_and_across_a_row_and_or_across_rows() -> None:
    cells = {
        "A1": "Tree", "B1": "Height", "C1": "Yield",
        "A2": "Apple", "B2": 18, "C2": 14,
        "A3": "Pear", "B3": 12, "C3": 10,
        "A4": "Apple", "B4": 9, "C4": 6,
        "E1": "Tree", "F1": "Height",
        "E2": "Apple", "F2": ">10",
        "E3": "Pear",
        "H1": "Tree", "H2": "App",
    }  # fmt: skip
    assert calculate('=DSUM(A1:C4,"Yield",E1:F2)', cells) == 14
    assert calculate('=DSUM(A1:C4,"Yield",E1:F3)', cells) == 24
    assert calculate("=DCOUNT(A1:C4,3,E1:F3)", cells) == 2
    assert calculate("=DCOUNTA(A1:C4,,E1:F3)", cells) == 2
    assert calculate('=DSUM(A1:C4,"Yield",H1:H2)', cells) == 20  # plain text matches a start
    assert calculate('=DGET(A1:C4,"Yield",E1:F2)', cells) == 14
    assert calculate('=DGET(A1:C4,"Yield",H1:H2)', cells) == "#NUM!"
    assert calculate('=DMAX(A1:C4,"Yield",E1:E3)', cells) == 14
    assert calculate('=DSUM(A1:C4,"Nothing",E1:F2)', cells) == "#VALUE!"


def test_computed_criteria_move_down_with_each_record() -> None:
    cells = {
        "A1": "N", "B1": "M", "A2": 1, "B2": 5, "A3": 2, "B3": 5, "A4": 3, "B4": 5,
        "D1": "Calc", "D2": "=A2>1",
    }  # fmt: skip
    assert calculate('=DSUM(A1:B4,"N",D1:D2)', cells) == 5


def test_complex_numbers_round_trip_as_text() -> None:
    assert calculate("=COMPLEX(3,4)") == "3+4i"
    assert calculate('=IMABS("3+4i")') == 5
    assert calculate('=IMPRODUCT("1+i","1-i")') == "2"
    assert calculate('=IMDIV("1","i")') == "-i"
    assert calculate('=IMSQRT("4")') == "2"
    assert calculate('=IMSUM("1+i","2j")') == "#VALUE!"
    assert calculate('=IMSUB("1+i","2j")') == "#NUM!"
    assert calculate('=IMREAL("3 + 4i")') == 3
    assert calculate('=COMPLEX(0,1,"j")') == "j"
    assert calculate('=IMSEC("1+i")') != "#NAME?"


@pytest.mark.parametrize(
    ("formula", "expected"),
    [
        ('=CONVERT(1,"mi","km")', 1.609344),
        ('=CONVERT(100,"C","F")', 212),
        ('=CONVERT(1,"kibyte","bit")', 8192),
        ('=CONVERT(1,"km2","m2")', 1_000_000),
        ('=CONVERT(2,"hr","mn")', 120),
        ('=CONVERT(1,"pc","c")', 1e-12),
    ],
)
def test_convert(formula: str, expected: float) -> None:
    assert calculate(formula) == pytest.approx(expected, rel=1e-12)


def test_convert_refuses_unknown_units_and_other_groups() -> None:
    assert calculate('=CONVERT(1,"m","kg")') == "#N/A"
    assert calculate('=CONVERT(1,"xx","m")') == "#N/A"
    assert calculate('=CONVERT(1,"ki","bit")') == "#N/A"
    assert calculate('=CONVERT(TRUE,"m","ft")') == "#VALUE!"


@pytest.mark.parametrize(
    ("formula", "expected"),
    [
        ("=BESSELJ(1,0)", 0.7651976865579666),
        ("=BESSELY(1,0)", 0.08825696421567697),
        ("=BESSELI(1,1)", 0.5651591039924851),
        ("=BESSELK(1,1)", 0.6019072301972346),
    ],
)
def test_bessel_functions(formula: str, expected: float) -> None:
    assert calculate(formula) == pytest.approx(expected, rel=1e-8)


def test_bessel_limits_follow_excel() -> None:
    assert calculate("=BESSELJ(134217728,0)") == "#NUM!"
    assert calculate("=BESSELI(709.8,0)") == "#NUM!"
    assert calculate("=BESSELK(0,0)") == "#NUM!"
    assert calculate("=BESSELJ(1,TRUE)") == "#VALUE!"


def test_urls_are_encoded_as_utf8() -> None:
    assert calculate('=ENCODEURL("a b&c~")') == "a%20b%26c%7E"
    assert calculate("=ENCODEURL(UNICHAR(233))") == "%C3%A9"


def test_securities_with_odd_periods_and_accruals() -> None:
    assert calculate("=ACCRINTM(DATE(2024,1,1),DATE(2024,7,1),0.06,1000,0)") == pytest.approx(30)
    assert calculate("=ACCRINT(DATE(2024,1,1),DATE(2024,7,1),DATE(2024,4,1),0.06,1000,2,0)") == (
        pytest.approx(15)
    )
    assert calculate(
        "=PRICEMAT(DATE(2024,6,1),DATE(2025,6,1),DATE(2024,1,1),0.05,0.06,0)"
    ) == pytest.approx(98.93867924528301)
    assert calculate("=PRICEMAT(DATE(2024,6,1),DATE(2025,6,1),DATE(2024,6,1),0.05,0.06)") == (
        "#NUM!"
    )
    assert calculate("=AMORLINC(2400,DATE(2023,12,31),DATE(2023,12,31),300,0,0.15)") == 360
    assert calculate("=AMORDEGRC(2400,DATE(2023,8,19),DATE(2023,12,31),300,1,0.6)") == "#NUM!"


def test_coupdays_on_an_actual_basis_follows_excel_at_month_ends() -> None:
    assert calculate("=COUPDAYS(DATE(2024,3,15),DATE(2024,9,30),2,1)") == 182
    assert calculate("=COUPDAYS(DATE(2024,3,15),DATE(2024,9,30),1,1)") == 366
    assert calculate("=COUPDAYS(DATE(2024,2,29),DATE(2027,8,31),4,1)") == 90
    assert calculate("=COUPDAYS(DATE(2023,12,31),DATE(2025,2,28),2,1)") == 184
    assert calculate("=COUPDAYS(DATE(2024,8,30),DATE(2027,8,31),2,1)") == 184
    assert calculate("=COUPDAYS(DATE(2024,8,31),DATE(2027,8,31),2,1)") == 181


def test_schedules_the_calculator_cannot_follow_are_left_uncalculated() -> None:
    with pytest.raises(UncalculableError):
        calculate("=ACCRINT(DATE(2024,1,15),DATE(2024,5,31),DATE(2024,6,30),0.06,1000,4)")
    with pytest.raises(UncalculableError):
        calculate("=COUPDAYS(DATE(2024,3,15),DATE(2025,1,30),4,1)")
    with pytest.raises(UncalculableError):
        calculate(
            "=ODDFPRICE(DATE(2024,1,5),DATE(2029,3,1),DATE(2022,11,11),DATE(2024,3,1),0.06,0.07,100,2,2)"
        )


def test_odd_coupon_prices_and_yields_agree() -> None:
    price = calculate(
        "=ODDFPRICE(DATE(2023,11,11),DATE(2029,3,1),DATE(2023,10,15),DATE(2024,3,1),0.0785,0.0625,100,2,1)"
    )
    back = calculate(
        f"=ODDFYIELD(DATE(2023,11,11),DATE(2029,3,1),DATE(2023,10,15),DATE(2024,3,1),0.0785,{price},100,2,1)"
    )
    assert back == pytest.approx(0.0625, rel=1e-8)
    last = calculate(
        "=ODDLPRICE(DATE(2024,2,7),DATE(2024,6,15),DATE(2023,10,15),0.0375,0.0405,100,2,0)"
    )
    assert last == pytest.approx(99.87828601472134, rel=1e-8)


def test_lambda_helpers_apply_a_lambda_to_rows_columns_and_elements() -> None:
    cells = {"A1": 1, "B1": 2, "A2": 3, "B2": 4}
    assert calculate("=INDEX(BYROW(A1:B2,LAMBDA(r,SUM(r))),2,1)", cells) == 7
    assert calculate("=INDEX(BYCOL(A1:B2,LAMBDA(c,MAX(c))),1,2)", cells) == 4
    assert calculate("=INDEX(MAP(A1:B2,LAMBDA(x,x*10)),2,2)", cells) == 40
    assert calculate("=REDUCE(0,A1:B2,LAMBDA(a,b,a+b))", cells) == 10
    assert calculate("=INDEX(SCAN(0,A1:B2,LAMBDA(a,b,a+b)),2,2)", cells) == 10
    assert calculate("=INDEX(MAKEARRAY(2,3,LAMBDA(r,c,r*c)),2,3)") == 6
    assert calculate("=INDEX(BYROW(A1:B2,SUM),1,1)", cells) == "#NAME?"
    assert calculate("=LET(f,LAMBDA(x,y,x+y),f(2,3))") == 5
    assert calculate("=LET(f,LAMBDA(x,x),f(1,2))") == "#VALUE!"
    assert calculate("=LET(f,LAMBDA(x,x),f)") == "#CALC!"
    assert calculate("=INDEX(MAP({1,2},{1,2,3},LAMBDA(a,b,a+b)),1,3)") == "#N/A"
    assert calculate("=INDEX(MAP({1,2,3},LAMBDA(x,{1,2})),1,1)") == "#CALC!"
