# ruff: noqa: E501
"""Database function cases (DSUM and the rest) for the Excel golden fixture."""

DB_ROWS: list[list[object]] = [
    ["Tree", "Height", "Age", "Yield", "Profit", "Flag"],
    ["Apple", 18, 20, 14, 105, True],
    ["Pear", 12, 12, 10, 96, False],
    ["Cherry", 13, 14, 9, 105, True],
    ["Apple", 14, 15, 10, 75, False],
    ["Pear", 9, 8, 8, 76.8, True],
    ["Apple", 8, 9, 6, 45, True],
    ["apple pie", 10, 5, None, 30, False],
    ["Plum", 11, 7, "n/a", 50, True],
    ["Peach", None, 6, 7, 60, False],
    ["*Star", 15, 10, 12, 80, True],
    ["Apple", 14, 15, 10, 75, False],
    ["Pear", "tall", 3, 4, "x", True],
]

# Criteria blocks: (header row, criteria rows). A "calc" header has formulas over the first record (row 2).
BLOCKS: dict[str, list[list[object]]] = {
    "apple": [["Tree"], ["Apple"]],
    "apple_tall": [["Tree", "Height"], ["Apple", ">10"]],
    "apple_or_pear": [["Tree"], ["Apple"], ["Pear"]],
    "mixed_rows": [["Tree", "Height"], ["Apple", ">10"], ["Pear", "<13"]],
    "range": [["Height", "Height"], [">10", "<15"]],
    "star": [["Tree"], ["A*"]],
    "equals": [["Tree"], ['="=Apple"']],
    "not": [["Tree"], ["<>Apple"]],
    "prefix": [["Tree"], ["App"]],
    "prefix_lower": [["Tree"], ["app"]],
    "empty_equals": [["Yield"], ['="="']],
    "blank_cell": [["Tree"], [None]],
    "number": [["Age"], [15]],
    "number_text": [["Age"], ['="=15"']],
    "ge": [["Age"], [">=14"]],
    "profit": [["Profit"], [">100"]],
    "question": [["Tree"], ["?pple"]],
    "tilde": [["Tree"], ["~*Star"]],
    "star_literal": [["Tree"], ["*Star"]],
    "header_only": [["Tree"]],
    "case_header": [["tree"], ["Apple"]],
    "unmatched": [["Tree"], ["Zebra"]],
    "trailing_blank": [["Tree"], ["Pear"], [None]],
    "bool": [["Flag"], [True]],
    "bool_text": [["Flag"], ["FALSE"]],
    "text_compare": [["Tree"], [">M"]],
    "num_on_text": [["Tree"], [5]],
    "text_on_num": [["Age"], ["a"]],
    "two_columns_or": [["Tree", "Age"], ["Apple", None], [None, 3]],
    "all_blank_row": [["Tree", "Age"], [None, None]],
    "foreign_header": [["Zzz"], [5]],
    "foreign_text": [["Zzz"], ["abc"]],
    "calc_gt": [[None], ["=DB!B2>10"]],
    "calc_label": [["Calc"], ["=DB!B2>10"]],
    "calc_and_plain": [["Tree", None], ["Apple", "=DB!C2>12"]],
    "calc_two_rows": [[None], ["=DB!B2>16"], ["=DB!D2=DB!D3"]],
    "calc_ratio": [[None], ["=DB!E2/DB!D2>9"]],
    "calc_text": [[None], ['=LEFT(DB!A2,1)="P"']],
    "calc_constant": [[None], [True]],
    "calc_number": [[None], ["=DB!D2"]],
    "calc_absolute": [[None], ["=DB!$D$2>10"]],
    "calc_error": [[None], ["=1/0"]],
    "yield_blank": [["Yield"], ["<>"]],
    "dup_tree": [["Tree", "Tree"], ["Apple", "Pear"]],
    "long_header": [
        ["Tree", "Height", "Age", "Yield", "Profit", "Flag"],
        ["Apple", ">10", None, None, None, None],
    ],
    "error_value": [["Yield"], ["#N/A"]],
    "pie": [["Tree"], ["apple pie"]],
    "plum": [["Tree"], ['="=Plum"']],
    "cherry": [["Tree"], ["Cherry"]],
    "peach": [["Tree"], ["Peach"]],
    "tall": [["Height"], ["tall"]],
    "text_yield": [["Yield"], ["n/a"]],
    "bool_text_two": [["Flag", "Tree"], [True, "P"]],
    "zero": [["Age"], [0]],
    "negative": [["Age"], ["<-1"]],
    "percent": [["Age"], ["1000%"]],
    "calc_false": [[None], [False]],
    "calc_formula_zero": [[None], ["=0*DB!B2"]],
    "calc_formula_text": [[None], ['="x"&DB!A2']],
    "calc_formula_true": [[None], ["=TRUE"]],
    "calc_formula_false": [[None], ["=FALSE"]],
    "text_number": [["Age"], ['="15"']],
    "text_number_prefix": [["Age"], ['="1"']],
    "text_height": [["Height"], ['="1"']],
    "percent_text": [["Age"], ['="10%"']],
    "decimal": [["Profit"], [76.8]],
    "decimal_text": [["Profit"], ['="76.8"']],
    "op_text": [["Profit"], ["<80"]],
    "le_bool": [["Flag"], ["<=TRUE"]],
    "space": [["Tree"], [" "]],
    "calc_zero": [[None], [0]],
    "calc_text_formula": [[None], ['="x"']],
    "calc_mixed": [["Tree", None, None], ["Apple", "=DB!B2>10", "=DB!D2<12"]],
    "calc_other_sheet": [[None], ["=Data!C2>5"]],
    "calc_relative_other": [[None], ["=DB!B2>Crit!Z1"]],
}

CRITERIA: list[list[object]] = []
BLOCK_RANGES: dict[str, str] = {}
for name, rows in BLOCKS.items():
    top = len(CRITERIA) + 1
    for row in rows:
        CRITERIA.append(row + [None] * (6 - len(row)))
    BLOCK_RANGES[name] = f"Crit!A{top}:{chr(64 + max(len(r) for r in rows))}{top + len(rows) - 1}"
    CRITERIA.append([None] * 6)

INPUTS: dict[str, list[list[object]]] = {"DB": DB_ROWS, "Crit": CRITERIA}
DATABASE = "DB!A1:F13"
FUNCTIONS = [
    "DSUM",
    "DAVERAGE",
    "DCOUNT",
    "DCOUNTA",
    "DGET",
    "DMAX",
    "DMIN",
    "DPRODUCT",
    "DSTDEV",
    "DSTDEVP",
    "DVAR",
    "DVARP",
]

CASES: list[str | tuple[str, float]] = []
for block, area in BLOCK_RANGES.items():
    for fn in FUNCTIONS:
        for field in (
            ('"Yield"', '"Profit"')
            if block != "apple"
            else ('"Yield"', '"Profit"', "4", '"Tree"', '"Height"', "6", '"flag"')
        ):
            CASES.append(f"={fn}({DATABASE},{field},{area})")
for fn in FUNCTIONS:
    area = BLOCK_RANGES["apple"]
    for field in (
        "0",
        "7",
        "-1",
        '"Zzz"',
        "TRUE",
        '""',
        "Data!Q3",
        "Data!G2",
        "2.9",
        '"2"',
        "NA()",
        "1/0",
        "{2}",
    ):
        CASES.append(f"={fn}({DATABASE},{field},{area})")
CASES += [
    f"=DCOUNT({DATABASE},,{BLOCK_RANGES[b]})"
    for b in ("apple", "apple_or_pear", "blank_cell", "unmatched", "header_only", "calc_gt")
]
CASES += [
    f"=DCOUNTA({DATABASE},,{BLOCK_RANGES[b]})"
    for b in ("apple", "apple_or_pear", "blank_cell", "unmatched", "header_only", "calc_gt")
]
CASES += [f'={fn}(DB!A1:F1,"Yield",{BLOCK_RANGES["apple"]})' for fn in FUNCTIONS]
CASES += [f'={fn}(DB!A1:F2,"Yield",{BLOCK_RANGES["apple"]})' for fn in FUNCTIONS]
CASES += [f"={fn}(DB!A1:A13,1,{BLOCK_RANGES['apple']})" for fn in ("DCOUNTA", "DGET", "DSUM")]
CASES += [
    f'=DSUM({DATABASE},"Yield",Crit!A1)',
    f'=DSUM({DATABASE},"Yield",5)',
    f'=DSUM({DATABASE},"Yield",Data!Z1:Z3)',
    '=DSUM(DB!A1:F13,"Yield",Crit!A1:A2)+1',
    f'=SUM(DSUM({DATABASE},"Yield",Crit!A1:A2),1)',
    '=DSUM(DB!A1:F13,"Yield",Crit!A1:A2)*2',
    f"=DSUM({DATABASE},C1,Crit!A1:A2)",
    f'=DGET({DATABASE},"Tree",{BLOCK_RANGES["star_literal"]})',
    f'=DGET({DATABASE},"Tree",{BLOCK_RANGES["tilde"]})',
    f'=DGET({DATABASE},"Profit",{BLOCK_RANGES["equals"]})',
]
CASES += [
    '=COUNTIF(DB!F2:F13,"<=TRUE")',
    '=COUNTIF(DB!F2:F13,">=TRUE")',
    '=COUNTIF(DB!F2:F13,"<TRUE")',
    '=COUNTIF(DB!F2:F13,">FALSE")',
    '=COUNTIF(DB!F2:F13,"<=FALSE")',
    '=COUNTIF(DB!A2:A13,"<=TRUE")',
    '=COUNTIF(DB!B2:B13,">=TRUE")',
]
