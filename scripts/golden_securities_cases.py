# ruff: noqa: E501
"""Security cases (accrued interest, odd coupons, French depreciation, coupon dates) for the golden fixture."""

import datetime as dt
import itertools

from excel_mcp.calc import functions  # noqa: F401  (registers the calculator's functions)
from excel_mcp.calc.registry import FUNCTIONS
from excel_mcp.calc.values import UncalculableError

INPUTS: dict[str, list[list[object]]] = {}
CASES: list[str | tuple[str, float]] = []
TOL = 1e-8


def month_end(date: dt.date) -> bool:
    return (date + dt.timedelta(days=1)).day == 1


def d(date: dt.date) -> str:
    return f"DATE({date.year},{date.month},{date.day})"


DATES = [
    dt.date(2023, 1, 31),
    dt.date(2023, 2, 28),
    dt.date(2023, 3, 31),
    dt.date(2023, 5, 15),
    dt.date(2023, 8, 31),
    dt.date(2023, 11, 30),
    dt.date(2024, 1, 15),
    dt.date(2024, 2, 29),
    dt.date(2024, 5, 31),
    dt.date(2024, 8, 31),
    dt.date(2024, 12, 31),
    dt.date(2025, 2, 28),
    dt.date(2025, 3, 15),
    dt.date(2026, 2, 28),
    dt.date(2027, 8, 31),
    dt.date(2028, 2, 29),
    dt.date(2028, 7, 4),
]
BASES = (0, 1, 2, 3, 4)

# ---- ACCRINTM, PRICEMAT, YIELDMAT
for issue, settlement in itertools.combinations(DATES[:12], 2):
    if (settlement - issue).days > 40:
        for basis in BASES:
            CASES.append((f"=ACCRINTM({d(issue)},{d(settlement)},0.06,1000,{basis})", TOL))
    CASES.append((f"=ACCRINTM({d(issue)},{d(settlement)},0.05,1000)", TOL))
for issue, settlement, maturity in itertools.combinations(DATES, 3):
    if (
        (maturity - settlement).days > 20
        and (settlement - issue).days > 20
        and (maturity.day + issue.day + settlement.day) % 11 == 0
    ):
        for basis in BASES:
            CASES.append(
                (f"=PRICEMAT({d(settlement)},{d(maturity)},{d(issue)},0.05,0.06,{basis})", TOL)
            )
            CASES.append(
                (f"=YIELDMAT({d(settlement)},{d(maturity)},{d(issue)},0.05,98.5,{basis})", TOL)
            )
        CASES.append((f"=PRICEMAT({d(settlement)},{d(maturity)},{d(issue)},0.05,0.06)", TOL))
        CASES.append((f"=YIELDMAT({d(settlement)},{d(maturity)},{d(issue)},0.05,98.5)", TOL))
CASES += [
    (f"=ACCRINTM({d(DATES[0])},{d(DATES[3])},{a},{b})", TOL)
    for a, b in ((0, 1000), (-0.05, 1000), (0.05, 0), (0.05, -1), (0.05, 100), (1e-10, 1000))
]
CASES += [
    "=ACCRINTM(DATE(2024,1,1),DATE(2024,1,1),0.05,1000)",
    "=ACCRINTM(DATE(2024,2,1),DATE(2024,1,1),0.05,1000)",
    "=ACCRINTM(DATE(2024,1,1),DATE(2024,6,1),0.05,1000,5)",
    "=ACCRINTM(DATE(2024,1,1),DATE(2024,6,1),0.05,1000,-1)",
    "=ACCRINTM(DATE(2024,1,1),DATE(2024,6,1),0.05,1000,1.9)",
    '=ACCRINTM("2024-01-01",DATE(2024,6,1),0.05,1000)',
    "=ACCRINTM(TRUE,DATE(2024,6,1),0.05,1000)",
    "=ACCRINTM(DATE(2024,1,1),DATE(2024,6,1),Data!G2,1000)",
    "=PRICEMAT(DATE(2024,6,1),DATE(2024,6,1),DATE(2024,1,1),0.05,0.06)",
    "=PRICEMAT(DATE(2024,6,1),DATE(2025,6,1),DATE(2024,6,1),0.05,0.06)",
    "=PRICEMAT(DATE(2024,6,1),DATE(2025,6,1),DATE(2024,7,1),0.05,0.06)",
    "=PRICEMAT(DATE(2024,6,1),DATE(2025,6,1),DATE(2024,1,1),-0.05,0.06)",
    "=PRICEMAT(DATE(2024,6,1),DATE(2025,6,1),DATE(2024,1,1),0,0.06)",
    "=PRICEMAT(DATE(2024,6,1),DATE(2025,6,1),DATE(2024,1,1),0.05,-0.06)",
    "=PRICEMAT(DATE(2024,6,1),DATE(2025,6,1),DATE(2024,1,1),0.05,0)",
    "=YIELDMAT(DATE(2024,6,1),DATE(2025,6,1),DATE(2024,1,1),0.05,0)",
    "=YIELDMAT(DATE(2024,6,1),DATE(2025,6,1),DATE(2024,1,1),0.05,-5)",
    "=YIELDMAT(DATE(2024,6,1),DATE(2025,6,1),DATE(2024,1,1),0,98)",
    "=YIELDMAT(DATE(2024,6,1),DATE(2025,6,1),DATE(2024,6,1),0.05,98)",
]

# ---- ACCRINT
for issue, first, settlement in itertools.combinations(DATES, 3):
    if (
        not month_end(first)
        and not any(x.month == 2 and month_end(x) for x in (issue, first, settlement))
        and (first - issue).days > 25
        and (settlement - issue).days > 25
        and (settlement - first).days < 400
        and (first - issue).days < 400
        and len(CASES) % 2 == 0
    ):
        for frequency in (1, 2, 4):
            for basis in (0, 1, 3):
                CASES.append(
                    (
                        f"=ACCRINT({d(issue)},{d(first)},{d(settlement)},0.06,1000,{frequency},{basis})",
                        TOL,
                    )
                )
        if not month_end(issue):
            CASES.append(
                (f"=ACCRINT({d(issue)},{d(first)},{d(settlement)},0.06,1000,2,0,FALSE)", TOL)
            )
        CASES.append((f"=ACCRINT({d(issue)},{d(first)},{d(settlement)},0.06,1000,4,2)", TOL))
        CASES.append((f"=ACCRINT({d(issue)},{d(first)},{d(settlement)},0.06,1000,1,4,TRUE)", TOL))
for issue, first, settlement in (
    (dt.date(2024, 1, 1), dt.date(2024, 7, 1), dt.date(2024, 4, 1)),
    (dt.date(2024, 1, 1), dt.date(2024, 7, 1), dt.date(2025, 3, 1)),
    (dt.date(2023, 11, 15), dt.date(2024, 7, 1), dt.date(2024, 9, 20)),
):
    for frequency, basis, method in itertools.product((1, 2, 4), BASES, ("TRUE", "FALSE")):
        if method == "FALSE" and (basis in (1, 2, 3) or month_end(issue)):
            continue
        CASES.append(
            (
                f"=ACCRINT({d(issue)},{d(first)},{d(settlement)},0.06,1000,{frequency},{basis},{method})",
                TOL,
            )
        )
CASES += [
    "=ACCRINT(DATE(2024,1,1),DATE(2024,7,1),DATE(2024,4,1),0.06,1000,3)",
    "=ACCRINT(DATE(2024,1,1),DATE(2024,7,1),DATE(2024,4,1),0.06,1000,0)",
    "=ACCRINT(DATE(2024,1,1),DATE(2024,7,1),DATE(2024,4,1),0,1000,2)",
    "=ACCRINT(DATE(2024,1,1),DATE(2024,7,1),DATE(2024,4,1),-0.06,1000,2)",
    "=ACCRINT(DATE(2024,1,1),DATE(2024,7,1),DATE(2024,4,1),0.06,0,2)",
    "=ACCRINT(DATE(2024,1,1),DATE(2024,7,1),DATE(2024,4,1),0.06,-1,2)",
    "=ACCRINT(DATE(2024,4,1),DATE(2024,7,1),DATE(2024,4,1),0.06,1000,2)",
    "=ACCRINT(DATE(2024,5,1),DATE(2024,7,1),DATE(2024,4,1),0.06,1000,2)",
    "=ACCRINT(DATE(2024,1,1),DATE(2023,7,1),DATE(2024,4,1),0.06,1000,2)",
    "=ACCRINT(DATE(2024,1,1),DATE(2024,1,1),DATE(2024,4,1),0.06,1000,2)",
    "=ACCRINT(DATE(2024,1,1),DATE(2024,7,1),DATE(2024,4,1),0.06,1000,2,5)",
    "=ACCRINT(DATE(2024,1,1),DATE(2024,7,1),DATE(2024,4,1),0.06,1000,2,0,2)",
    '=ACCRINT(DATE(2024,1,1),DATE(2024,7,1),DATE(2024,4,1),0.06,1000,2,0,"x")',
    "=ACCRINT(DATE(2024,1,1),DATE(2024,7,1),DATE(2024,4,1),0.06,1000,2,0,Data!Q3)",
    "=ACCRINT(DATE(2024,1,1),DATE(2024,7,1),DATE(2024,4,1),0.06,1000,2,Data!Q3)",
    "=ACCRINT(DATE(2024,1,1),DATE(2024,7,1),DATE(2024,4,1),0.06,Data!Q3,2)",
]

# ---- AMORLINC and AMORDEGRC
for fn in ("AMORLINC", "AMORDEGRC"):
    for purchased, first in (
        (dt.date(2023, 8, 19), dt.date(2023, 12, 31)),
        (dt.date(2024, 2, 28), dt.date(2024, 12, 31)),
        (dt.date(2023, 1, 1), dt.date(2023, 12, 31)),
        (dt.date(2023, 12, 31), dt.date(2023, 12, 31)),
        (dt.date(2023, 6, 15), dt.date(2024, 3, 31)),
    ):
        for cost, salvage, rate in (
            (2400, 300, 0.15),
            (10000, 0, 0.2),
            (1000, 100, 0.5),
            (50000, 5000, 0.08),
            (2400, 2400, 0.1),
            (2400, 300, 0.3),
            (2400.5, 300.25, 0.17),
        ):
            for basis in (0, 1, 3, 4):
                for period in (0, 1, 2, 3, 5, 8, 11):
                    if (purchased.day + cost + period + rate + basis) % 5 < 2:
                        CASES.append(
                            (
                                f"={fn}({cost},{d(purchased)},{d(first)},{salvage},{period},{rate},{basis})",
                                TOL,
                            )
                        )
    CASES += [
        (f"={fn}(2400,DATE(2023,8,19),DATE(2023,12,31),300,{p},0.15,1)", TOL) for p in range(12)
    ]
    CASES += [(f"={fn}(2400,DATE(2023,8,19),DATE(2023,12,31),300,{p},0.6)", TOL) for p in range(4)]
    CASES += [(f"={fn}(2400,DATE(2023,8,19),DATE(2023,12,31),300,{p},0.3)", TOL) for p in range(7)]
    CASES += [
        (f"={fn}(2400,DATE(2023,8,19),DATE(2023,12,31),300,{p},0.22,0)", TOL) for p in range(7)
    ]
    CASES += [
        f"={fn}(2400,DATE(2023,8,19),DATE(2023,12,31),300,1,0.15,2)",
        f"={fn}(0,DATE(2023,8,19),DATE(2023,12,31),300,1,0.15)",
        f"={fn}(-5,DATE(2023,8,19),DATE(2023,12,31),300,1,0.15)",
        f"={fn}(2400,DATE(2023,8,19),DATE(2023,12,31),-1,1,0.15)",
        f"={fn}(2400,DATE(2023,8,19),DATE(2023,12,31),3000,1,0.15)",
        f"={fn}(2400,DATE(2023,8,19),DATE(2023,12,31),300,-1,0.15)",
        f"={fn}(2400,DATE(2023,8,19),DATE(2023,12,31),300,1,0)",
        f"={fn}(2400,DATE(2023,8,19),DATE(2023,12,31),300,1,-0.1)",
        f"={fn}(2400,DATE(2023,12,31),DATE(2023,8,19),300,1,0.15)",
        f"={fn}(2400,DATE(2023,8,19),DATE(2023,12,31),300,1.9,0.15)",
        f"={fn}(2400,DATE(2023,8,19),DATE(2023,12,31),300,1,0.15,5)",
        f"={fn}(2400,DATE(2023,8,19),DATE(2023,12,31),300,1,0.15,-1)",
        f"={fn}(2400,DATE(2023,8,19),DATE(2023,12,31),300,1,2)",
        f"={fn}(2400,DATE(2023,8,19),DATE(2023,12,31),300,1,1E-5)",
        f"={fn}(1E10,DATE(2023,8,19),DATE(2023,12,31),0,3,0.15)",
        f"={fn}(2400,DATE(2023,8,19),Data!G2,300,1,0.15)",
    ]


def calculable(settlement: dt.date, maturity: dt.date, frequency: int) -> bool:
    """Whether the calculator answers COUPDAYS (a few settlements at month-ends it declines)."""
    serial = lambda day: (day - dt.date(1899, 12, 30)).days  # noqa: E731
    try:
        FUNCTIONS["COUPDAYS"].call(serial(settlement), serial(maturity), frequency, 1)
    except UncalculableError:
        return False
    return True


# ---- coupon dates around month-ends
MATURITIES = [
    dt.date(2025, 2, 28),
    dt.date(2024, 2, 29),
    dt.date(2024, 9, 30),
    dt.date(2027, 8, 31),
    dt.date(2026, 7, 31),
    dt.date(2025, 4, 30),
    dt.date(2025, 3, 29),
    dt.date(2025, 5, 30),
    dt.date(2025, 1, 31),
    dt.date(2024, 12, 31),
]
SETTLEMENTS = [
    dt.date(2023, 12, 31),
    dt.date(2024, 1, 31),
    dt.date(2024, 2, 28),
    dt.date(2024, 2, 29),
    dt.date(2024, 3, 15),
    dt.date(2024, 3, 31),
    dt.date(2024, 5, 31),
    dt.date(2024, 6, 30),
    dt.date(2024, 8, 30),
    dt.date(2024, 8, 31),
    dt.date(2024, 11, 30),
    dt.date(2023, 8, 31),
    dt.date(2023, 2, 28),
]
for maturity in MATURITIES:
    for settlement in SETTLEMENTS:
        if settlement >= maturity:
            continue
        for frequency in (1, 2, 4):
            for basis in BASES:
                for name in (
                    "COUPDAYS",
                    "COUPDAYBS",
                    "COUPDAYSNC",
                    "COUPNCD",
                    "COUPPCD",
                    "COUPNUM",
                ):
                    if (
                        name == "COUPDAYS"
                        and basis == 1
                        and not calculable(settlement, maturity, frequency)
                    ):
                        continue
                    if (maturity.day + settlement.day + frequency + basis + len(name)) % 7 == 0:
                        CASES.append(f"={name}({d(settlement)},{d(maturity)},{frequency},{basis})")
                if (maturity.day + settlement.day + frequency + basis) % 7 == 0:
                    CASES.append(
                        (
                            f"=PRICE({d(settlement)},{d(maturity)},0.05,0.06,100,{frequency},{basis})",
                            TOL,
                        )
                    )
                    CASES.append(
                        (
                            f"=YIELD({d(settlement)},{d(maturity)},0.05,98.5,100,{frequency},{basis})",
                            TOL,
                        )
                    )
                    CASES.append(
                        (
                            f"=DURATION({d(settlement)},{d(maturity)},0.05,0.06,{frequency},{basis})",
                            TOL,
                        )
                    )


# ---- odd first and last coupon periods
def months(date: dt.date, count: int) -> dt.date:
    index = date.year * 12 + date.month - 1 + count
    year, month = divmod(index, 12)
    last = (dt.date(year + (month + 1) // 12, (month + 1) % 12 + 1, 1) - dt.timedelta(days=1)).day
    return dt.date(year, month + 1, min(date.day, last))


ODD_FIRST = [
    (dt.date(2023, 11, 11), dt.date(2024, 3, 1)),
    (dt.date(2023, 1, 15), dt.date(2024, 3, 1)),
    (dt.date(2023, 12, 20), dt.date(2024, 2, 29)),
    (dt.date(2023, 8, 31), dt.date(2024, 2, 29)),
    (dt.date(2022, 6, 30), dt.date(2024, 5, 31)),
    (dt.date(2023, 3, 5), dt.date(2023, 9, 15)),
    (dt.date(2024, 1, 31), dt.date(2024, 7, 31)),
]
for issue, first in ODD_FIRST:
    if month_end(first):
        continue
    for frequency in (1, 2, 4):
        step = 12 // frequency
        maturity = months(first, 5 * step)
        for settlement in (
            issue + dt.timedelta(days=1),
            issue + (first - issue) / 2,
            first - dt.timedelta(days=3),
        ):
            settlement = dt.date(settlement.year, settlement.month, settlement.day)
            for basis in BASES:
                long_first = issue < months(first, -12 // frequency)
                if basis in (2, 3) and long_first:
                    continue
                if (issue.day + frequency + basis + settlement.day) % 2 == 0:
                    CASES.append(
                        (
                            f"=ODDFPRICE({d(settlement)},{d(maturity)},{d(issue)},{d(first)},0.06,0.07,100,{frequency},{basis})",
                            TOL,
                        )
                    )
                    CASES.append(
                        (
                            f"=ODDFYIELD({d(settlement)},{d(maturity)},{d(issue)},{d(first)},0.06,98.5,100,{frequency},{basis})",
                            TOL,
                        )
                    )
ODD_LAST = [
    dt.date(2023, 8, 15),
    dt.date(2023, 11, 30),
    dt.date(2024, 2, 29),
    dt.date(2023, 12, 31),
    dt.date(2024, 1, 31),
]
for last in ODD_LAST:
    if month_end(last):
        continue
    for frequency in (1, 2, 4):
        for gap in (40, 100, 200, 300, 420):
            maturity = last + dt.timedelta(days=gap)
            for settlement in (
                last + dt.timedelta(days=3),
                last + dt.timedelta(days=gap // 2),
                maturity - dt.timedelta(days=7),
            ):
                for basis in BASES:
                    if (last.day + frequency + basis + gap) % 2 == 0:
                        CASES.append(
                            (
                                f"=ODDLPRICE({d(settlement)},{d(maturity)},{d(last)},0.06,0.07,100,{frequency},{basis})",
                                TOL,
                            )
                        )
                        CASES.append(
                            (
                                f"=ODDLYIELD({d(settlement)},{d(maturity)},{d(last)},0.06,98.5,100,{frequency},{basis})",
                                TOL,
                            )
                        )
CASES += [
    "=ODDFPRICE(DATE(2023,11,11),DATE(2029,3,1),DATE(2023,10,15),DATE(2024,3,1),0.0785,0.0625,100,2,1)",
    "=ODDFYIELD(DATE(2023,11,11),DATE(2029,3,1),DATE(2023,10,15),DATE(2024,3,1),0.0575,84.5,100,2,0)",
    "=ODDLPRICE(DATE(2024,2,7),DATE(2024,6,15),DATE(2023,10,15),0.0375,0.0405,100,2,0)",
    "=ODDLYIELD(DATE(2024,4,20),DATE(2024,6,15),DATE(2023,12,24),0.0375,99.875,100,2,0)",
    "=ODDFPRICE(DATE(2024,1,1),DATE(2029,3,1),DATE(2023,10,15),DATE(2024,3,1),0.06,0.07,100,3)",
    "=ODDFPRICE(DATE(2024,1,1),DATE(2029,3,1),DATE(2023,10,15),DATE(2024,3,1),0.06,0.07,100,0)",
    "=ODDFPRICE(DATE(2024,1,1),DATE(2029,3,1),DATE(2023,10,15),DATE(2024,3,1),-0.06,0.07,100,2)",
    "=ODDFPRICE(DATE(2024,1,1),DATE(2029,3,1),DATE(2023,10,15),DATE(2024,3,1),0,0.07,100,2)",
    "=ODDFPRICE(DATE(2024,1,1),DATE(2029,3,1),DATE(2023,10,15),DATE(2024,3,1),0.06,-0.07,100,2)",
    "=ODDFPRICE(DATE(2024,1,1),DATE(2029,3,1),DATE(2023,10,15),DATE(2024,3,1),0.06,0,100,2)",
    "=ODDFPRICE(DATE(2024,1,1),DATE(2029,3,1),DATE(2023,10,15),DATE(2024,3,1),0.06,0.07,0,2)",
    "=ODDFPRICE(DATE(2024,1,1),DATE(2029,3,1),DATE(2023,10,15),DATE(2024,3,1),0.06,0.07,100,2,5)",
    "=ODDFPRICE(DATE(2024,1,1),DATE(2029,3,1),DATE(2024,1,1),DATE(2024,3,1),0.06,0.07,100,2)",
    "=ODDFPRICE(DATE(2023,10,15),DATE(2029,3,1),DATE(2023,10,15),DATE(2024,3,1),0.06,0.07,100,2)",
    "=ODDFPRICE(DATE(2024,1,1),DATE(2029,3,1),DATE(2023,10,15),DATE(2024,1,1),0.06,0.07,100,2)",
    "=ODDFPRICE(DATE(2024,3,1),DATE(2029,3,1),DATE(2023,10,15),DATE(2024,3,1),0.06,0.07,100,2)",
    "=ODDFPRICE(DATE(2024,1,1),DATE(2024,3,1),DATE(2023,10,15),DATE(2024,3,1),0.06,0.07,100,2)",
    "=ODDFPRICE(DATE(2024,1,1),DATE(2029,4,1),DATE(2023,10,15),DATE(2024,3,1),0.06,0.07,100,2)",
    "=ODDFPRICE(DATE(2024,1,1),DATE(2029,3,15),DATE(2023,10,15),DATE(2024,3,1),0.06,0.07,100,2)",
    "=ODDFYIELD(DATE(2024,1,1),DATE(2029,3,1),DATE(2023,10,15),DATE(2024,3,1),0.06,0,100,2)",
    "=ODDFYIELD(DATE(2024,1,1),DATE(2029,3,1),DATE(2023,10,15),DATE(2024,3,1),0.06,-5,100,2)",
    (
        "=ODDFYIELD(DATE(2024,1,1),DATE(2029,3,1),DATE(2023,10,15),DATE(2024,3,1),0.06,1E-5,100,2)",
        TOL,
    ),
    "=ODDLPRICE(DATE(2024,2,7),DATE(2024,6,15),DATE(2024,2,7),0.0375,0.0405,100,2)",
    "=ODDLPRICE(DATE(2024,6,15),DATE(2024,6,15),DATE(2023,10,15),0.0375,0.0405,100,2)",
    "=ODDLPRICE(DATE(2024,7,1),DATE(2024,6,15),DATE(2023,10,15),0.0375,0.0405,100,2)",
    "=ODDLPRICE(DATE(2024,2,7),DATE(2024,6,15),DATE(2023,10,15),-0.0375,0.0405,100,2)",
    "=ODDLPRICE(DATE(2024,2,7),DATE(2024,6,15),DATE(2023,10,15),0.0375,-0.0405,100,2)",
    "=ODDLPRICE(DATE(2024,2,7),DATE(2024,6,15),DATE(2023,10,15),0.0375,0.0405,0,2)",
    "=ODDLPRICE(DATE(2024,2,7),DATE(2024,6,15),DATE(2023,10,15),0.0375,0.0405,100,5)",
    "=ODDLPRICE(DATE(2024,2,7),DATE(2024,6,15),DATE(2023,10,15),0.0375,0.0405,100,2,7)",
    "=ODDLYIELD(DATE(2024,2,7),DATE(2024,6,15),DATE(2023,10,15),0.0375,0,100,2)",
    "=ODDLYIELD(DATE(2024,2,7),DATE(2024,6,15),DATE(2023,10,15),0.0375,-1,100,2)",
    "=ODDLYIELD(DATE(2024,2,7),DATE(2024,6,15),DATE(2023,10,15),0,99,100,2)",
]


# COUPDAYS on an actual basis with month-end maturities: settlements every 9 days and the days
# around the plain and the month-end coupon dates, where Excel's period length changes.
MONTH_END_MATURITIES = [
    dt.date(2027, 8, 31),
    dt.date(2027, 5, 31),
    dt.date(2027, 4, 30),
    dt.date(2027, 2, 28),
    dt.date(2026, 7, 31),
]
for maturity in MONTH_END_MATURITIES:
    for frequency in (1, 2, 4):
        around = {dt.date(2023, 6, 1) + dt.timedelta(days=9 * i) for i in range(75)}
        for year in (2023, 2024, 2025):
            for month in range(1, 13):
                for day in (27, 28, 29, 30, 31):
                    try:
                        around.add(dt.date(year, month, day))
                    except ValueError:
                        continue
        for settlement in sorted(around):
            if (
                dt.date(2023, 6, 1) <= settlement < maturity
                and (settlement.day + settlement.month + frequency) % 3 == 0
            ) and calculable(settlement, maturity, frequency):
                CASES.append(f"=COUPDAYS({d(settlement)},{d(maturity)},{frequency},1)")
