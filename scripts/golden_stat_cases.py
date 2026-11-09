"""Statistics and regression cases for the Excel golden fixture; see golden_cases.CASES."""

INPUTS: dict[str, list[list[object]]] = {
    "Stat": [
        # A    B       C     D       E     F
        [79, "x", 70, 1, 2, 2],
        [85, 2, 79, None, 2, 3],
        [78, None, 89, True, 3, 3],
        [85, 4.5, 85, "5", 3, 4],
        [50, True, None, 2, 4, 5],
        [81, 85, "t", 3, 4, 6],
        [95, 9, 70, 3, 5, 7],
        [88, None, 79, "abc", 5, 8],
        [97, 7, 89, 4, 6, 9],
        [None, 3, None, 4, 6, 10],
    ],
}

A_RANGES = ["Stat!A1:A10", "Stat!B1:B10", "Stat!D1:D10", "Data!F2:F11", "Data!Q2:Q6", "Stat!C1:C10"]
CASES: list[str | tuple[str, float]] = []
for name in (
    "STDEVA",
    "STDEVPA",
    "VARA",
    "VARPA",
    "AVERAGEA",
    "MAXA",
    "MINA",
    "SKEW.P",
    "DEVSQ",
    "AVEDEV",
    "STDEV.S",
    "STDEV.P",
    "VAR.S",
    "VAR.P",
    "SKEW",
    "KURT",
):
    CASES += [f"={name}({r})" for r in A_RANGES]
    CASES += [
        f"={name}(Stat!A1:A3,Stat!B1:B3)",
        f"={name}(1)",
        f"={name}(1,2)",
        f"={name}(1,2,3)",
        f'={name}("5",2,3)',
        f'={name}("a",2,3)',
        f"={name}(TRUE,2,4)",
        f"={name}(1,2,Data!G2)",
        f"={name}(Data!Z1:Z3)",
        f"={name}({{1,2,4,8}})",
        f'={name}({{1,TRUE,"a",4}})',
        f"={name}(Stat!A1:A10,Stat!D1:D10,5)",
        f"={name}(Stat!E1:E10)",
        f"={name}(Stat!E1)",
        f"={name}(Stat!E1,Stat!E2)",
    ]
CASES += [
    "=SKEW.P(5,5,5)",
    "=SKEW.P(1,2,3)",
    "=SKEW.P(1,1,9)",
    "=STDEVA(5)",
    "=STDEVPA(5)",
    "=VARPA(5)",
    "=VARA(5)",
    "=VARA(TRUE,FALSE)",
    "=STDEVA(TRUE,FALSE,TRUE)",
]
CASES += [
    f"=MODE.MULT({r})"
    for r in (
        "Stat!E1:E10",
        "Stat!F1:F10",
        "Stat!A1:A10",
        "Stat!C1:C10",
        "Stat!B1:B10",
        "Data!Q2:Q6",
    )
]
CASES += [
    f"=INDEX(MODE.MULT({r}),{k})"
    for r in ("Stat!E1:E10", "Stat!F1:F10", "Stat!A1:A10", "Stat!C1:C10")
    for k in (1, 2, 3)
]
CASES += [
    "=ROWS(MODE.MULT(Stat!E1:E10))",
    "=ROWS(MODE.MULT(Stat!F1:F10))",
    "=COLUMNS(MODE.MULT(Stat!E1:E10))",
    "=ROWS(MODE.MULT(1,1,2,2,3))",
    "=INDEX(MODE.MULT({1,1,2,2,3,3,3,4,4,4}),2)",
    "=ROWS(MODE.MULT({1,1,2,2,3,3,3,4,4,4}))",
    "=MODE.MULT(Stat!E1:E2,Stat!F1:F3)",
]
for p in (0, 0.1, 0.2, 0.25, 0.5, 0.9, 0.99, 1, -0.1, 1.1, "TRUE", '"0.2"'):
    CASES += [
        f"=TRIMMEAN(Stat!A1:A10,{p})",
        f"=TRIMMEAN(Stat!F1:F10,{p})",
        f"=TRIMMEAN(Data!Q2:Q6,{p})",
    ]
CASES += [
    "=TRIMMEAN(Stat!A1:A3,Data!G2)",
    "=TRIMMEAN(Data!Z1:Z3,0.2)",
    "=TRIMMEAN({1,2,3,4,5,6,7,8,9,10},0.3)",
    "=TRIMMEAN({1,2,3,4,5,6,7,8,9,10},0.19)",
    "=TRIMMEAN({1,2,3},0.9)",
]
CASES += [
    f"=STEYX({y},{x})"
    for y, x in (
        ("Dist!A4:A13", "Dist!B4:B13"),
        ("Stat!A1:A10", "Stat!D1:D10"),
        ("Data!M2:M11", "Data!N2:N11"),
        ("Dist!A4:A5", "Dist!B4:B5"),
        ("Dist!A4:A6", "Dist!B4:B6"),
        ("Dist!A4:A13", "Dist!B4:B12"),
        ("{1,2,3,4}", "{3,3,3,3}"),
        ("{1,2,3,4}", "{2,4,6,8}"),
        ("{1,2,3,5}", "{2,4,6,8}"),
        ("Stat!A1:A3", "Stat!B1:B3"),
    )
]
CASES += [
    f"=INDEX(FREQUENCY({d},{b}),{k})"
    for d in ("Stat!A1:A10", "Stat!B1:B10", "Stat!E1:E10")
    for b in ("Stat!C1:C10", "Stat!F1:F10", "{70,80,90}", "{90,70,80}", "{80,80,80}", "Stat!C5:C6")
    for k in (1, 2, 3, 4, 5)
]
CASES += [
    f"=ROWS(FREQUENCY({d},{b}))"
    for d in ("Stat!A1:A10", "Stat!B1:B10")
    for b in (
        "Stat!C1:C10",
        "{70,80,90}",
        "{90,70,80}",
        "Stat!C5:C6",
        "Stat!B1:B1",
        "Data!Z1:Z3",
        "5",
        "{5}",
    )
]
CASES += [
    "=COLUMNS(FREQUENCY(Stat!A1:A10,{70,80,90}))",
    "=INDEX(FREQUENCY(Stat!A1:A10,{70,80,90}),1,2)",
    "=SUM(FREQUENCY(Stat!A1:A10,{70,80,90}))",
    "=INDEX(FREQUENCY(Stat!A1:A10,Data!G2:G3),1)",
    "=INDEX(FREQUENCY({1,2,3},{1,2,3}),4)",
    "=INDEX(FREQUENCY({1,2,3},{1.5,2.5}),2)",
    "=INDEX(FREQUENCY(5,{1,10}),2)",
    "=INDEX(FREQUENCY(5,{1,10}),3)",
]
CASES += [
    f"=ECMA.CEILING({a},{b})"
    for a, b in (
        (2.5, 1),
        (2.5, 0.5),
        (-2.5, 1),
        (-2.5, -1),
        (2.5, -1),
        (0, 1),
        (2.5, 0),
        (4.42, 0.05),
        (-4.42, 0.05),
        (1.5, 2),
        ('"2.5"', 1),
    )
]
CASES += [
    f"={f}({a})"
    for f in ("AVERAGEA", "MAXA", "MINA", "STDEVA", "VARPA", "COUNTA")
    for a in (
        "{1,TRUE}",
        '{1,"a"}',
        '{"a"}',
        "{TRUE}",
        "{1,TRUE,4}",
        '{1,"a",4}',
        "{-2,TRUE,-4}",
        '{-2,"a",-4}',
        "{-2,FALSE,-4}",
        '{1,4,"a",5}',
        "{TRUE,TRUE}",
        '{"a","b"}',
        "({1,2,3}>1)",
        "({1,2,3}>1)*1",
        "{1,4,TRUE},Stat!B1:B3",
    )
]
CASES += [
    "=INDEX(FREQUENCY({-1,0,1},Data!Z1:Z3),1)",
    "=INDEX(FREQUENCY({-1,0,1},Data!Z1:Z3),2)",
    '=INDEX(FREQUENCY({-1,0,1},{"a"}),1)',
    "=INDEX(FREQUENCY(Data!G2:G4,{1,2}),1)",
    "=INDEX(FREQUENCY({1,2,3},Data!G2:G4),1)",
    "=ROWS(FREQUENCY({1,2,3},Stat!C1:C10))",
    '=INDEX(FREQUENCY({1,2,3},{2,"a",TRUE}),1)',
    '=ROWS(FREQUENCY({1,2,3},{2,"a",TRUE}))',
    "=INDEX(FREQUENCY(Data!Z1:Z3,{1,2}),3)",
    "=INDEX(FREQUENCY({1,2,3,4},{3,1,3,2}),1)",
    "=INDEX(FREQUENCY({1,2,3,4},{3,1,3,2}),2)",
    "=INDEX(FREQUENCY({1,2,3,4},{3,1,3,2}),3)",
    "=INDEX(FREQUENCY({1,2,3,4},{3,1,3,2}),4)",
    "=INDEX(FREQUENCY({1,2,3,4},{3,1,3,2}),5)",
    "=INDEX(FREQUENCY({1,1.5,2.0000000000000004},{2,1}),2)",
]
