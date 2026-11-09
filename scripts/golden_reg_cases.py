# ruff: noqa: E501
"""Regression cases (LINEST, LOGEST, TREND, GROWTH) for the Excel golden fixture."""

Y = [1.5, 3.1, 4.4, 6.2, 7.1, 9.3, 10.2, 12.4]
X1 = [1, 2, 3, 4, 5, 6, 7, 8]
X2 = [2.1, 1.9, 4.2, 3.8, 6.1, 5.8, 8.3, 7.9]
X3 = [2 * x for x in X1]
NOISE = [0.3, -0.7, 0.1, 0.9, -0.4, 0.6, -0.2, 0.5]
X4 = [x + 1e-6 * n for x, n in zip(X1, NOISE, strict=True)]
X5 = [x + 1e-9 * n for x, n in zip(X1, NOISE, strict=True)]
X6 = [x + 1e-12 * n for x, n in zip(X1, NOISE, strict=True)]
X7 = [x + 1e-14 * n for x, n in zip(X1, NOISE, strict=True)]
CONST = [5] * 8
GROW = [2.1, 3.0, 4.6, 6.9, 10.1, 15.2, 22.0, 33.5]
EXACT = [1 + 2 * x for x in X1]
EXACT2 = [1 + 2 * a + 3 * b for a, b in zip(X1, [1, 0, 2, 1, 3, 0, 1, 2], strict=True)]
X8 = [1, 0, 2, 1, 3, 0, 1, 2]
TEXTY = [1.5, 3.1, "x", 6.2, 7.1, 9.3, 10.2, 12.4]
COLUMNS = [Y, X1, X2, X3, X4, X5, X6, X7, CONST, GROW, EXACT, EXACT2, X8, TEXTY]
# columns: A=y B=x1 C=x2 D=2*x1 E..H = x1 with noise 1e-6, 1e-9, 1e-12, 1e-14, I=const, J=growth y, K=exact y, L=exact 2 vars y, M=x8, N=text in y

INPUTS: dict[str, list[list[object]]] = {
    "Reg": [list(row) for row in zip(*COLUMNS, strict=True)]
    + [[None] * 14, ["y", *Y], ["x1", *X1], ["x2", *X2], ["g", *GROW]],
}

CASES: list[str | tuple[str, float]] = []
TOL = 1e-9


def rows(top: int, bottom: int, left: str, right: str | None = None) -> str:
    return f"Reg!{left}{top}:{right or left}{bottom}"


Y_ = "Reg!A1:A8"
G_ = "Reg!J1:J8"
K_ = "Reg!K1:K8"
L_ = "Reg!L1:L8"


def element_cases(
    function: str,
    y: str,
    x: str,
    columns: int,
    const: str,
    stats: str = "TRUE",
    tolerance: float = TOL,
) -> None:
    args = f"{y},{x},{const},{stats}" if const else f"{y},{x}"
    last_row = 5 if stats == "TRUE" else 1
    for r in range(1, last_row + 1):
        for c in range(1, columns + 2):
            CASES.append((f"=INDEX({function}({args}),{r},{c})", tolerance))
    CASES.append((f"=ROWS({function}({args}))", tolerance))
    CASES.append((f"=COLUMNS({function}({args}))", tolerance))


X_SETS = [
    ("Reg!B1:B8", 1),
    ("Reg!B1:C8", 2),
    ("Reg!B1:D8", 3),
    ("Reg!C1:D8", 2),
    ("Reg!B1:E8", 4),
    ("Reg!B1:B8,Reg!I1:I8", 0),
]
for fn, y in (("LINEST", Y_), ("LOGEST", G_)):
    for x, k in X_SETS[:5]:
        for const in ("TRUE", "FALSE"):
            element_cases(fn, y, x, k, const)
    element_cases(fn, y, "Reg!B1:B8", 1, "")
    element_cases(fn, y, "Reg!B1:B8", 1, "TRUE", "FALSE")
    element_cases(fn, y, "Reg!B1:B8", 1, "FALSE", "FALSE")
    element_cases(fn, y, "Reg!B1:B8", 1, "1", "0")
    element_cases(fn, y, "Reg!B1:B8", 1, '"TRUE"', "TRUE")
    element_cases(fn, y, "Reg!I1:I8", 1, "TRUE", "FALSE")  # its statistics are garbage in Excel
    element_cases(fn, y, "Reg!I1:I8", 1, "FALSE")
    element_cases(fn, y, "Reg!B1:I8", 8, "TRUE", "FALSE")
    for column in "EFGH":
        element_cases(fn, y, "Reg!B1:B8", 1, "TRUE")
        element_cases(fn, y, f"Reg!{column}1:{column}8", 1, "TRUE")
    CASES += [
        (f"=ROWS({fn}({y}))", TOL),
        (f"=COLUMNS({fn}({y}))", TOL),
        (f"=INDEX({fn}({y}),1,1)", TOL),
        (f"=INDEX({fn}({y},,TRUE,TRUE),3,1)", TOL),
        (f"=INDEX({fn}({y},,FALSE,TRUE),3,1)", TOL),
    ]
    CASES += [
        (f"=INDEX({fn}({y},Reg!B1:B8,,TRUE),3,1)", TOL),
        (f"=INDEX({fn}({y},Reg!B1:B8,Data!Q3,TRUE),3,1)", TOL),
        (f"=INDEX({fn}({y},Reg!B1:B8,TRUE,Data!Q3),3,1)", TOL),
        (f"=INDEX({fn}({y},Reg!B1:B8,FALSE,Data!Q3),1,2)", TOL),
    ]
    # shapes and errors
    CASES += [
        f"=INDEX({fn}(Reg!A1:B8,Reg!C1:C8),1,1)",
        f"=INDEX({fn}({y},Reg!B1:B7),1,1)",
        f"=INDEX({fn}({y},Reg!B1:B9),1,1)",
        f"=INDEX({fn}({y},Reg!B1:C7),1,1)",
        f"=INDEX({fn}({y},Reg!B1:B2),1,1)",
        f"=INDEX({fn}(Reg!N1:N8,Reg!B1:B8),1,1)",
        f"=INDEX({fn}({y},Reg!N1:N8),1,1)",
        f"=INDEX({fn}(Reg!Z1:Z8,Reg!B1:B8),1,1)",
        f"=INDEX({fn}({y},Data!G2:G9),1,1)",
        f"=INDEX({fn}(Data!G2:G9,Reg!B1:B8),1,1)",
        f"=INDEX({fn}(Reg!A1:A1,Reg!B1:B1),1,1)",
        f"=INDEX({fn}(Reg!A1:A2,Reg!B1:B2),1,1)",
        f"=INDEX({fn}(Reg!A1:A2,Reg!B1:B2,TRUE,TRUE),1,1)",
        f"=INDEX({fn}(Reg!A1:A2,Reg!B1:B2,TRUE,TRUE),3,1)",
        f"=INDEX({fn}(Reg!A1:A2,Reg!B1:B2,TRUE,TRUE),4,2)",
        f"=INDEX({fn}(Reg!A1:A3,Reg!B1:C3,TRUE,TRUE),1,1)",
        f"=INDEX({fn}(Reg!A1:A3,Reg!B1:C3,TRUE,TRUE),2,1)",
        f"=INDEX({fn}(Reg!A1:A3,Reg!B1:C3,TRUE,TRUE),3,1)",
        f"=INDEX({fn}(Reg!A1:A3,Reg!B1:C3,TRUE,TRUE),4,1)",
        f"=INDEX({fn}(Reg!A1:A3,Reg!B1:C3,TRUE,TRUE),4,2)",
        f"=INDEX({fn}(Reg!A1:A3,Reg!B1:B3,FALSE,TRUE),4,2)",
        f"=INDEX({fn}(Reg!A1:A2,Reg!B1:B2,FALSE,TRUE),4,2)",
        f"=INDEX({fn}(Reg!A1:A2,Reg!B1:C2,FALSE,TRUE),1,1)",
    ]
    # row layout
    CASES += [
        (f"=INDEX({fn}(Reg!B10:I10,Reg!B11:I11,TRUE,TRUE),{r},{c})", TOL)
        for r in range(1, 6)
        for c in (1, 2)
    ]
    CASES += [
        (f"=INDEX({fn}(Reg!B10:I10,Reg!B11:I12,TRUE,TRUE),{r},{c})", TOL)
        for r in range(1, 6)
        for c in (1, 2, 3)
    ]
    CASES += [
        (f"=INDEX({fn}(Reg!B10:I10,Reg!B11:B12),1,1)", TOL),
        (f"=INDEX({fn}(Reg!B10:I10,Reg!B11:I11),1,1)", TOL),
        (f"=INDEX({fn}(Reg!B10:I10,Reg!B11:C11),1,1)", TOL),
        (f"=INDEX({fn}(Reg!A1:A8,Reg!B11:I11),1,1)", TOL),
        (f"=INDEX({fn}(Reg!B10:I10),1,1)", TOL),
    ]
    # exact fits
    CASES += [
        (
            f"=INDEX({fn}({K_ if fn == 'LINEST' else '{2,4,8,16,32,64,128,256}'},Reg!B1:B8,TRUE,TRUE),{r},{c})",
            1e-8,
        )
        for r in range(1, 6)
        for c in (1, 2)
        if (r, c) != (4, 1)  # F of a perfect fit is rounding noise
    ]
    CASES += [
        (f"=INDEX({fn}({L_},Reg!B1:B8,TRUE,TRUE),{r},{c})", 1e-8)
        for r in range(1, 6)
        for c in (1, 2)
        if (r, c) != (4, 1)
    ]

# TREND and GROWTH
for fn, y in (("TREND", Y_), ("GROWTH", G_)):
    for x, new in (
        ("Reg!B1:B8", None),
        ("Reg!B1:B8", "Reg!B1:B8"),
        ("Reg!B1:B8", "{9;10}"),
        ("Reg!B1:B8", "{9,10}"),
        ("Reg!B1:B8", "{1,2;3,4}"),
        ("Reg!B1:C8", None),
        ("Reg!B1:C8", "Reg!B1:C3"),
        ("Reg!B1:C8", "{9,9}"),
        ("Reg!B1:C8", "{9;9}"),
        ("Reg!B1:C8", "{9,9,9}"),
        ("Reg!B1:D8", "Reg!B1:D3"),
        ("Reg!B1:D8", None),
        ("Reg!B1:E8", "Reg!B1:E3"),
        (None, None),
        (None, "{9;10}"),
        (None, "{9,10}"),
    ):
        for const in (None, "TRUE", "FALSE"):
            parts = [y, x or "", new or "", const or ""]
            while parts and parts[-1] == "":
                parts.pop()
            args = ",".join(parts)
            CASES += [
                (f"=INDEX({fn}({args}),{r},{c})", TOL) for r in (1, 2, 3, 4, 8) for c in (1, 2)
            ]
            CASES += [(f"=ROWS({fn}({args}))", TOL), (f"=COLUMNS({fn}({args}))", TOL)]
    CASES += [(f"=INDEX({fn}(Reg!B10:I10,Reg!B11:I12),1,{c})", TOL) for c in (1, 2, 8)]
    CASES += [(f"=INDEX({fn}(Reg!B10:I10,Reg!B11:I12,Reg!B11:D12),1,{c})", TOL) for c in (1, 2, 3)]
    CASES += [
        (f"=ROWS({fn}(Reg!B10:I10,Reg!B11:I12,Reg!B11:D12))", TOL),
        (f"=COLUMNS({fn}(Reg!B10:I10,Reg!B11:I12,Reg!B11:D12))", TOL),
        (f"=ROWS({fn}(Reg!B10:I10,Reg!B11:I11,Reg!B11:D11))", TOL),
        (f"=COLUMNS({fn}(Reg!B10:I10,Reg!B11:I11,Reg!B11:D11))", TOL),
    ]
    CASES += [
        (f"=INDEX({fn}({y},Reg!B1:B8,Reg!B1:B3,TRUE),1,1)", TOL),
        f"=INDEX({fn}({y},Reg!B1:B7),1,1)",
        f"=INDEX({fn}({y},Reg!B1:B8,Reg!Z1:Z3),1,1)",
        f"=INDEX({fn}({y},Reg!B1:B8,Data!G2:G3),1,1)",
        f"=INDEX({fn}({y},Reg!B1:B8,{{1,TRUE}}),1,1)",
        f'=INDEX({fn}({y},Reg!B1:B8,{{"a"}}),1,1)',
        f"=INDEX({fn}(Reg!N1:N8,Reg!B1:B8),1,1)",
        f"=INDEX({fn}(Reg!A1:B8,Reg!C1:C8),1,1)",
    ]
CASES += [
    ("=INDEX(GROWTH(Reg!A1:A8,Reg!B1:B8),1,1)", TOL),
    "=INDEX(GROWTH(Reg!N1:N8,Reg!B1:B8),1,1)",
    ("=INDEX(GROWTH({1,2,0,4},{1,2,3,4}),1,1)", TOL),
    ("=INDEX(GROWTH({1,2,-1,4},{1,2,3,4}),1,1)", TOL),
    ("=INDEX(LOGEST({1,2,0,4},{1,2,3,4}),1,1)", TOL),
    ("=INDEX(LOGEST({1,2,-1,4},{1,2,3,4}),1,1)", TOL),
]

LEVELS = [
    1e-2,
    1e-3,
    1e-4,
    1e-5,
    1e-6,
    1e-7,
    1e-8,
    1e-9,
    1e-10,
    1e-11,
    1e-12,
    1e-13,
    1e-14,
    *(k * 1e-6 for k in (3.1, 3.2, 3.3, 3.4, 3.5, 3.6, 3.7, 3.8, 3.9)),
]
TOL_COLUMNS: list[list[float]] = []
for scale, shift in ((1, 0),):
    for level in LEVELS:
        TOL_COLUMNS.append([scale * (x + shift) for x in X1])
        TOL_COLUMNS.append(
            [scale * (x + shift + level * n) for x, n in zip(X1, NOISE, strict=True)]
        )
INPUTS["Tol"] = [list(row) for row in zip(*TOL_COLUMNS, strict=True)]
