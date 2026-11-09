# ruff: noqa: E501
"""Distribution cases for the Excel golden fixture; see golden_cases.CASES."""

ITERATIVE = 1e-9

INPUTS: dict[str, list[list[object]]] = {
    "Dist": [
        # A      B      C   D   E   F   G    H    I     J     K
        [10, 20, 30, None, 12, 18, 28, None, None, None, None],
        [15, 25, 35, None, 13, 27, 37, None, None, None, None],
        [None] * 11,
        [4.1, 3.9, None, 10, None, 1, 0.1, 0.1, 0.5, "x", True],
        [5.3, 5.9, None, 12, None, 2, 0.2, 0.2, 0.5, 3, None],
        [6.2, 6.0, None, 9, None, 3, 0.3, 0.3, None, None, None],
        [7.8, 8.2, None, 14, None, 4, 0.4, 0.5, None, None, None],
        [5.5, 5.1, None, 11, None, None, None, None, None, None, None],
        [6.9, 7.3, None, 13, None, None, None, None, None, None, None],
        [4.8, 4.4, None, 10.5, None, None, None, None, None, None, None],
        [7.1, 7.7, None, None, None, None, None, None, None, None, None],
        [6.0, 6.4, None, None, None, None, None, None, None, None, None],
        [5.2, 4.9, None, None, None, None, None, None, None, None, None],
    ],
}

SAMPLE_A, SAMPLE_B = "Dist!A4:A13", "Dist!B4:B13"
SHORT = "Dist!D4:D9"


def calls(name: str, *argument_sets: tuple[object, ...]) -> list[str]:
    return [f"={name}({','.join(str(a) for a in args)})" for args in argument_sets]


def iterative(cases: list[str]) -> list[tuple[str, float]]:
    """Excel's inverse functions are only accurate to about 1e-7 for 1e9 or more degrees of freedom."""
    huge = ("999999999", "10000000000")
    return [(c, 1e-6 if "INV" in c and any(h in c for h in huge) else ITERATIVE) for c in cases]


NUMBERS = [(0.5, 2, 3), (0.2, 1.5, 4), (0.9, 5, 1), (0, 2, 3), (1, 2, 3), (1.5, 2, 3), (-0.1, 2, 3)]
ERRORS = [
    ("Dist!J4", 2, 3),
    ("Dist!K4", 2, 3),
    ("Dist!H5", 2, 3),
    ("Dist!I3", 2, 3),
    ("Dist!H2", 2, 3),
]

CASES: list[str | tuple[str, float]] = []

# beta
for cumulative in ("TRUE", "FALSE"):
    CASES += iterative(
        calls(
            "BETA.DIST", *[(*n, cumulative) for n in NUMBERS], *[(*e, cumulative) for e in ERRORS]
        )
    )
    CASES += iterative(
        calls(
            "BETA.DIST",
            (2, 2, 3, cumulative, 1, 3),
            (0.5, 2, 3, cumulative, 1, 3),
            (2, 2, 3, cumulative, 3, 1),
            (2, 2, 3, cumulative, 2, 2),
            (0.5, 0.5, 0.5, cumulative),
            (0, 1, 1, cumulative),
            (1, 1, 1, cumulative),
            (0, 0.5, 2, cumulative),
            (1, 2, 0.5, cumulative),
            (0, 2, 2, cumulative),
            (1, 2, 2, cumulative),
            (0.3, 50, 80, cumulative),
            (0.3, 0, 2, cumulative),
            (0.3, -1, 2, cumulative),
            ('"0.4"', 2, 3, cumulative),
            ("TRUE", 2, 3, cumulative),
            (0.4, 2, 3, cumulative, 0, 2),
        )
    )
CASES += iterative(
    calls(
        "BETA.INV",
        (0.5, 2, 3),
        (0.1, 2, 3),
        (0.99, 2, 3),
        (0.5, 0.5, 0.5),
        (0.001, 5, 5),
        (0.5, 100, 200),
        (0, 2, 3),
        (1, 2, 3),
        (1.1, 2, 3),
        (-0.1, 2, 3),
        (0.5, 0, 3),
        (0.5, 2, -1),
        (0.5, 2, 3, 1, 5),
        (0.5, 2, 3, 5, 1),
        (0.5, 2, 3, 2, 2),
        (0.3, 1, 1),
        (0.999999, 2, 3),
        (1e-8, 2, 3),
        ('"0.5"', 2, 3),
        ("Dist!J4", 2, 3),
    )
)
CASES += iterative(
    calls(
        "BETADIST",
        (0.5, 2, 3),
        (0.2, 1.5, 4),
        (2, 2, 3, 1, 3),
        (0.5, 2, 3, 1, 3),
        (1.5, 2, 3),
        (0.5, 0, 3),
    )
)
CASES += iterative(calls("BETAINV", (0.5, 2, 3), (0.1, 2, 3, 1, 5), (0, 2, 3), (1, 2, 3)))

# gamma
for cumulative in ("TRUE", "FALSE"):
    CASES += iterative(
        calls(
            "GAMMA.DIST",
            (2, 3, 2, cumulative),
            (0.5, 0.5, 1, cumulative),
            (0, 2, 1, cumulative),
            (0, 1, 2, cumulative),
            (0, 0.5, 1, cumulative),
            (10, 3, 2, cumulative),
            (100, 3, 2, cumulative),
            (-1, 3, 2, cumulative),
            (2, 0, 2, cumulative),
            (2, 3, 0, cumulative),
            (2, 100, 1, cumulative),
            (150, 100, 1, cumulative),
            ('"2"', 3, 2, cumulative),
            ("Dist!J4", 3, 2, cumulative),
            ("Dist!H2", 3, 2, cumulative),
        )
    )
CASES += iterative(
    calls(
        "GAMMA.INV",
        (0.5, 3, 2),
        (0.01, 3, 2),
        (0.99, 3, 2),
        (0, 3, 2),
        (1, 3, 2),
        (1.1, 3, 2),
        (0.5, 0.2, 1),
        (0.5, 100, 1),
        (1e-6, 3, 1),
        (0.9999, 3, 1),
        (0.5, 0, 1),
        (0.5, 3, 0),
        (0.5, 3, -1),
    )
)
CASES += iterative(calls("GAMMADIST", (2, 3, 2, 1), (2, 3, 2, 0)))
CASES += iterative(calls("GAMMAINV", (0.5, 3, 2), (0, 3, 2)))

# chi-square
for df in (1, 2, 3, 10, 2.7, 0, 0.5, -1, 100, 1000, '"4"', "Dist!J4"):
    CASES += iterative(
        calls(
            "CHISQ.DIST",
            (3.5, df, "TRUE"),
            (3.5, df, "FALSE"),
            (0, df, "TRUE"),
            (0, df, "FALSE"),
            (-1, df, "TRUE"),
        )
    )
    CASES += iterative(calls("CHISQ.DIST.RT", (3.5, df), (0, df), (-1, df), (200, df)))
    CASES += iterative(
        calls("CHISQ.INV", (0.5, df), (0, df), (1, df), (0.95, df), (1e-6, df), (1.5, df))
    )
    CASES += iterative(calls("CHISQ.INV.RT", (0.5, df), (0, df), (1, df), (0.05, df), (1.5, df)))
    CASES += iterative(calls("CHIDIST", (3.5, df)) + calls("CHIINV", (0.05, df)))
CASES += iterative(
    calls(
        "CHISQ.DIST",
        (3, 10000000001, "TRUE"),
        (3, 10000000000, "TRUE"),
        (3, 9999999999, "TRUE"),
        (3, 1e10, "TRUE"),
    )
)
CASES += iterative(
    calls(
        "F.DIST",
        (2.5, 10000000001, 5, "TRUE"),
        (2.5, 5, 10000000001, "TRUE"),
        (2.5, 5, 10000000000.9, "TRUE"),
    )
    + calls("T.DIST", (1, 10000000001, "TRUE"))
)
CASES += iterative(calls("CHISQ.INV.RT", (0.5, 9999999999), (0.5, 10000000000)))
CASES += [
    ("=CHISQ.INV.RT(1E-09,1)", 1e-7),
    ("=CHISQ.INV.RT(1E-09,2)", 1e-7),
]  # Excel's own result is off by 4e-9
CASES += iterative(
    [
        "=CHISQ.TEST(Dist!A1:C2,Dist!E1:G2)",
        "=CHISQ.TEST(Dist!A1:C1,Dist!E1:G1)",
        "=CHISQ.TEST(Dist!A1:A2,Dist!E1:E2)",
        "=CHISQ.TEST(Dist!A1:B2,Dist!E1:G2)",
        "=CHISQ.TEST(Dist!A1,Dist!E1)",
        "=CHISQ.TEST(Dist!A1:C2,Dist!E1:G2*0)",
        "=CHITEST(Dist!A1:C2,Dist!E1:G2)",
        "=CHISQ.TEST(Dist!A1:C2,Dist!A1:C2)",
        "=CHISQ.TEST(Dist!A1:C2,Dist!D1:F2)",
        "=CHISQ.TEST(Dist!A1:C3,Dist!E1:G3)",
        "=CHISQ.TEST(Dist!A1:C2,Dist!E1:G2*1)",
    ]
)

# F
for d in (
    (3, 5),
    (1, 1),
    (2, 2),
    (10, 20),
    (2.9, 5.9),
    (0, 5),
    (5, 0),
    (-1, 5),
    (5, 10000000000),
    (5, 9999999999),
    ("Dist!J4", 5),
    (50, 1000),
):
    CASES += iterative(
        calls(
            "F.DIST",
            (2.5, *d, "TRUE"),
            (2.5, *d, "FALSE"),
            (0, *d, "TRUE"),
            (0, *d, "FALSE"),
            (-1, *d, "TRUE"),
            (100, *d, "FALSE"),
        )
    )
    CASES += iterative(calls("F.DIST.RT", (2.5, *d), (0, *d), (-1, *d), (100, *d)))
    CASES += iterative(
        calls("F.INV", (0.5, *d), (0, *d), (1, *d), (0.95, *d), (1e-6, *d), (1.5, *d))
    )
    CASES += iterative(
        calls("F.INV.RT", (0.5, *d), (0, *d), (1, *d), (0.05, *d), (1e-6, *d), (1.5, *d))
    )
    CASES += iterative(calls("FDIST", (2.5, *d)) + calls("FINV", (0.05, *d)))
CASES += iterative(
    [
        f"=F.TEST({SAMPLE_A},{SAMPLE_B})",
        f"=F.TEST({SAMPLE_A},{SHORT})",
        f"=FTEST({SAMPLE_A},{SAMPLE_B})",
        "=F.TEST(Dist!A4,Dist!B4)",
        f"=F.TEST({SAMPLE_A},Dist!C4:C13)",
        f"=F.TEST({SAMPLE_A},Dist!A4:A13*2)",
        "=F.TEST({1,2,3,4},{2,4,6,9})",
        "=F.TEST({1,1,1},{2,4,6})",
        f"=F.TEST({SAMPLE_A},Dist!A4:A5)",
        f"=F.TEST(Dist!J4:J6,{SAMPLE_B})",
    ]
)

# t
for df in (
    1,
    2,
    5,
    30,
    2.7,
    0,
    0.9,
    -1,
    1000,
    10000000000,
    9999999999,
    '"5"',
    "Dist!J4",
    "Dist!H2",
):
    CASES += iterative(
        calls(
            "T.DIST",
            (1.5, df, "TRUE"),
            (1.5, df, "FALSE"),
            (-1.5, df, "TRUE"),
            (-1.5, df, "FALSE"),
            (0, df, "TRUE"),
            (0, df, "FALSE"),
            (40, df, "TRUE"),
            (-40, df, "TRUE"),
        )
    )
    CASES += iterative(calls("T.DIST.2T", (1.5, df), (0, df), (-1.5, df), (40, df)))
    CASES += iterative(calls("T.DIST.RT", (1.5, df), (-1.5, df), (0, df), (40, df)))
    CASES += iterative(
        calls("T.INV", (0.9, df), (0.1, df), (0.5, df), (0, df), (1, df), (1e-8, df), (1.5, df))
    )
    CASES += iterative(
        calls("T.INV.2T", (0.1, df), (1, df), (0.5, df), (0, df), (1e-8, df), (1.5, df))
    )
    CASES += iterative(
        calls(
            "TDIST",
            (1.5, df, 1),
            (1.5, df, 2),
            (-1, df, 1),
            (1.5, df, 3),
            (1.5, df, 0),
            (1.5, df, 1.9),
        )
        + calls("TINV", (0.1, df))
    )
CASES += iterative(calls("T.INV.2T", (0.05, 1), (0.05, 2), (0.05, 3), (0.001, 1), (0.99, 1)))
CASES += iterative(
    calls(
        "CONFIDENCE.T",
        (0.05, 2, 10),
        (0.05, 2, 2),
        (0.05, 2, 1),
        (0.05, 2, 0),
        (0.05, 0, 10),
        (0, 2, 10),
        (1, 2, 10),
        (0.05, 2, 10.9),
        (0.05, -1, 10),
        (0.01, 15, 500),
    )
)
CASES += iterative(
    calls("CONFIDENCE.NORM", (0.05, 2, 10), (0.05, 2, 1), (0.05, 2, 0), (0.05, 0, 10), (1, 2, 10))
)

TTEST_ARGS = [
    (SAMPLE_A, SAMPLE_B),
    (SAMPLE_A, SHORT),
    ("Dist!A4:A8", "Dist!B4:B8"),
    ("Dist!A4:A13", "Dist!A4:A13"),
    ("Dist!A4:A5", "Dist!B4:B5"),
    ("Dist!A4", "Dist!B4"),
    ("Dist!J4:J6", SAMPLE_B),
]
for first, second in TTEST_ARGS:
    for tails in (1, 2, 3, 0, 1.9):
        for kind in (1, 2, 3, 4, 0):
            if (kind in (1, 2, 3) and tails in (1, 2)) or (kind, tails) in ((4, 2), (1, 3)):
                CASES += iterative([f"=T.TEST({first},{second},{tails},{kind})"])
CASES += iterative(
    [
        f"=TTEST({SAMPLE_A},{SAMPLE_B},2,3)",
        "=T.TEST({1,2,3,4},{2,4,6,9},2,1)",
        "=T.TEST({1,1,1},{1,1,1},2,2)",
        "=T.TEST({1,1,1},{1,1,1},2,3)",
        "=T.TEST({1,1,1},{2,2,2},2,3)",
        "=T.TEST({1,2,3},{2,3,4},2,1)",
    ]
)
CASES += iterative(
    [
        f"=Z.TEST({SAMPLE_A},5)",
        f"=Z.TEST({SAMPLE_A},5,1.2)",
        f"=ZTEST({SAMPLE_A},6)",
        f"=Z.TEST({SAMPLE_A},5,0)",
        f"=Z.TEST({SAMPLE_A},5,-1)",
        "=Z.TEST(Dist!A4,5)",
        "=Z.TEST(Dist!A4,5,1)",
        "=Z.TEST({1,2,3,4,5},3)",
        "=Z.TEST(Dist!J4:J5,3)",
        "=Z.TEST(Dist!K9:K10,3)",
        f"=Z.TEST({SAMPLE_A},Dist!J4)",
        f"=Z.TEST({SAMPLE_A},10)",
        f"=Z.TEST({SAMPLE_A},2)",
        f"=Z.TEST({SAMPLE_A},5,Dist!J4)",
    ]
)

# lognormal, Weibull
for sd in (1, 0.5, 0, -1, '"1"', "Dist!J4"):
    CASES += iterative(
        calls(
            "LOGNORM.DIST",
            (2, 0.5, sd, "TRUE"),
            (2, 0.5, sd, "FALSE"),
            (0, 0.5, sd, "TRUE"),
            (-1, 0.5, sd, "TRUE"),
            (1e-5, 0.5, sd, "FALSE"),
        )
    )
    CASES += iterative(
        calls(
            "LOGNORM.INV",
            (0.3, 0.5, sd),
            (0, 0.5, sd),
            (1, 0.5, sd),
            (0.999, 0.5, sd),
            (1e-7, 0.5, sd),
        )
    )
    CASES += iterative(calls("LOGNORMDIST", (2, 0.5, sd)) + calls("LOGINV", (0.3, 0.5, sd)))
for a, b in ((2, 3), (1, 3), (0.5, 3), (0, 3), (2, 0), (-1, 3), ('"2"', 3), ("Dist!J4", 3)):
    CASES += iterative(
        calls(
            "WEIBULL.DIST",
            (2, a, b, "TRUE"),
            (2, a, b, "FALSE"),
            (0, a, b, "TRUE"),
            (0, a, b, "FALSE"),
            (-1, a, b, "TRUE"),
            (50, a, b, "FALSE"),
        )
    )
    CASES += iterative(calls("WEIBULL", (2, a, b, 1)))

# discrete
HYPER = [
    (1, 4, 8, 20),
    (0, 4, 8, 20),
    (4, 4, 8, 20),
    (5, 4, 8, 20),
    (3, 4, 8, 20),
    (1, 4, 8, 6),
    (2, 4, 18, 20),
    (1, 4, 18, 20),
    (-1, 4, 8, 20),
    (1, 25, 8, 20),
    (1, 4, 25, 20),
    (1.9, 4.9, 8.9, 20.9),
    (1, 0, 8, 20),
    (0, 0, 0, 1),
    (1, 4, 8, 0),
    (10, 40, 80, 200),
    ("Dist!J4", 4, 8, 20),
]
for h in HYPER:
    CASES += iterative(calls("HYPGEOM.DIST", (*h, "TRUE"), (*h, "FALSE")) + calls("HYPGEOMDIST", h))
CASES += iterative(
    calls("HYPGEOM.DIST", (50, 100, 300, 600, "TRUE"), (500, 1000, 3000, 6000, "FALSE"))
)
for n in (
    (0, 2, 0.5),
    (3, 2, 0.5),
    (3, 2, 1),
    (3, 2, 0),
    (0, 1, 1),
    (0, 1, 0),
    (-1, 2, 0.5),
    (3, 0, 0.5),
    (3, 2, 1.1),
    (3, 2, -0.1),
    (3.9, 2.9, 0.5),
    (10, 5, 0.3),
    (50, 20, 0.4),
    ("Dist!J4", 2, 0.5),
):
    CASES += iterative(
        calls("NEGBINOM.DIST", (*n, "TRUE"), (*n, "FALSE")) + calls("NEGBINOMDIST", n)
    )
for b in (
    (10, 0.5, 0.5),
    (10, 0.5, 0),
    (10, 0.5, 1),
    (10, 0.5, 0.999999),
    (0, 0.5, 0.5),
    (-1, 0.5, 0.5),
    (10, 1.5, 0.5),
    (10, 0.5, 1.5),
    (10.9, 0.5, 0.3),
    (100, 0.3, 0.95),
    (1000, 0.01, 0.5),
    (10, 0, 0.5),
    (10, 1, 0.5),
    (10, 0.5, 0.0001),
    ("Dist!J4", 0.5, 0.5),
):
    CASES += iterative(calls("BINOM.INV", b) + calls("CRITBINOM", b))
for r in (
    (10, 0.5, 3, 5),
    (10, 0.5, 3),
    (10, 0.5, 0, 10),
    (10, 0.5, 5, 3),
    (10, 0.5, 11),
    (10, 0.5, -1),
    (10, 0.5, 3, 11),
    (-1, 0.5, 0),
    (10, 1.1, 3),
    (10, 0.5, 3.9, 5.9),
    (100, 0.3, 20, 40),
    (10, 0, 0, 0),
    (10, 1, 10, 10),
    ("Dist!J4", 0.5, 3),
    (10, 0.5, 3, "Dist!H2"),
):
    CASES += iterative(calls("BINOM.DIST.RANGE", r))
CASES += iterative(
    [
        "=PROB(Dist!F4:F7,Dist!G4:G7,2)",
        "=PROB(Dist!F4:F7,Dist!G4:G7,2,3)",
        "=PROB(Dist!F4:F7,Dist!G4:G7,3,2)",
        "=PROB(Dist!F4:F7,Dist!G4:G7,0,10)",
        "=PROB(Dist!F4:F7,Dist!G4:G7,2.5)",
        "=PROB(Dist!F4:F7,Dist!G4:G7,2.5,3.5)",
        "=PROB(Dist!F4:F7,Dist!H4:H7,2)",
        "=PROB(Dist!F4:F7,Dist!G4:G6,2)",
        "=PROB(Dist!F4:F7,Dist!I4:I7,2)",
        "=PROB(Dist!F4:F7,{0.1,0.2,0.3,0.4},2,4)",
        "=PROB({1,2,3},{0.2,0.3,0.5},2)",
        "=PROB({1,2,3},{0.2,0.3,0.5000001},2)",
        "=PROB({1,2,3},{0.2,0.3,0.4999999},2)",
        "=PROB({1,2,3},{0.2,0.3,0.5},Dist!J4)",
        "=PROB({1,2,3},{0.2,0.3,0.5},2,Dist!J5)",
        "=PROB({1,2,3},{1.5,-0.5,0},2)",
        "=PROB({1,2,3},{2,-1,0},1)",
        '=PROB({1,2,3},{"a",0.5,0.5},2)',
        "=PROB({1,2,3},{0.5,0.5,1E-10},2)",
        "=PROB({1,2,3},{0.5,0.5,1E-8},2)",
        "=PROB({1,2,3},{0.5,0.5,-1E-12},2)",
        "=PROB({1,2,3},{0.5,0.5,1E-11},2)",
        "=PROB({1,2,3},{0.5,0.5,1E-12},2)",
        "=PROB({1,2,3},{0.5,0.5,1E-13},2)",
        "=PROB({1,2,3},{0.5,0.5,1E-14},2)",
        "=PROB({1,2,3},{0.5,0.5,1E-15},2)",
        "=PROB({1,2,3},{0.5,0.5,3E-16},2)",
        "=PROB({1,2,3},{0.5,0.5,2E-15},2)",
        "=PROB({1,2,3},{0.5,0.5,3E-15},2)",
        "=PROB({1,2,3},{0.5,0.5,5E-15},2)",
        "=PROB({1,2,3},{0.5,0.5,7E-15},2)",
        "=PROB({1,2,3},{0.5,0.5,9E-15},2)",
        "=PROB({1,2,3},{0.5,0.5,-2E-15},2)",
        "=PROB({1,2,3},{0.5,0.5,-5E-15},2)",
        "=PROB({1,2,3},{0.5,0.5,-1E-15},2)",
        "=PROB({1,2,3},{0.5,0.5,-1E-14},2)",
        "=PROB({1,2,3},{0.2,0.3,0},2)",
        "=PROB({1,2,3},{1,0,0},1)",
        "=PROB({1,2,3},{1,0,0},1,3)",
        "=PROB({1,2,3},{-0.5,1,0.5},2)",
        "=PROB({1,2,3},{0.2,0.3,0.5},TRUE)",
    ]
)
STEPS = [
    *(f"{k}*2^-52" for k in (*range(1, 31), -1, -2)),
    *(f"{k}*2^-53" for k in (-1, -2, -3, -4, -5, -6, -7, -8, -9)),
]
INPUTS["Probs"] = [[1, 2, 3], *([0.5, 0.5, f"={step}"] for step in STEPS)]
CASES += [f"=PROB(Probs!A1:C1,Probs!A{r}:C{r},2)" for r in range(2, len(STEPS) + 2)]
CASES += [
    "=PROB({1,2,3},{0.7,0.2,0.1},2)",
    "=PROB({1,2,3},{0.1,0.2,0.7},2)",
    "=PROB({1,2,3,4},{0.1,0.2,0.3,0.4},2)",
    "=PROB({1,2,3,4,5,6,7,8,9,10},{0.1,0.1,0.1,0.1,0.1,0.1,0.1,0.1,0.1,0.1},2)",
    "=PROB({1,2,3},{0.3,0.3,0.4},2)",
    "=PROB({1,2,3},{0.9,0.05,0.05},2)",
    "=PROB({1,2,3,4,5},{0.2,0.2,0.2,0.2,0.2},2)",
]
