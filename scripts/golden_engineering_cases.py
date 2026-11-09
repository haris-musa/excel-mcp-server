# ruff: noqa: E501
"""Engineering cases (complex numbers, CONVERT, Bessel functions) for the Excel golden fixture."""

COMPLEX_TEXTS = [
    "3+4i",
    "3-4i",
    "-3+4i",
    "-3-4i",
    "4i",
    "-4i",
    "i",
    "-i",
    "+i",
    "3",
    "-3",
    "0",
    "1+i",
    "1-i",
    "2.5+1.5i",
    "1E2+3i",
    "1e-3-2e-3j",
    "3+4j",
    "3+4I",
    "3+4J",
    "3+4 i",
    " 3+4i",
    "3+4",
    "3+",
    "+",
    "-",
    "",
    "abc",
    "3i+4",
    "3+4ij",
    "3+4i+5",
    "1++i",
    "0+0i",
    "00.5+0.5i",
    "1+4.i",
    ".5+.5i",
    "3+4*i",
    "i3",
    "3 + 4i",
    "1e",
    "1E+",
    "1+i1",
    "3+-4i",
    "5i7",
    "j",
    "-j",
    "1+j",
    "0i",
    "0j",
    "-0",
    "1E-20+1E-20i",
    "1E200+1E200i",
    "123456789.123456789+1i",
    "0.1+0.2i",
    "1e5i",
    "2+0i",
    "0+2i",
    "+3+4i",
    "+3",
    "3e",
    "3+4k",
    "k",
    "3+4i ",
    "1,5+2i",
    "1.5.5+i",
    "5-i",
    "5+-i",
]
INPUT_CELLS = [
    ("Data!Q3", "blank"),
    ("Data!G2", "error"),
    ("Data!P2", "TRUE cell"),
    ("Data!C2", "number cell"),
]
INPUTS: dict[str, list[list[object]]] = {}
CASES: list[str | tuple[str, float]] = []


def quoted(text: str) -> str:
    return '"' + text.replace('"', '""') + '"'


UNARY = [
    "IMREAL",
    "IMAGINARY",
    "IMABS",
    "IMARGUMENT",
    "IMCONJUGATE",
    "IMEXP",
    "IMSQRT",
    "IMLN",
    "IMLOG10",
    "IMLOG2",
    "IMSIN",
    "IMCOS",
    "IMTAN",
    "IMCOT",
    "IMSEC",
    "IMCSC",
    "IMSINH",
    "IMCOSH",
    "IMSECH",
    "IMCSCH",
]
TRIG = {
    "IMEXP",
    "IMSIN",
    "IMCOS",
    "IMTAN",
    "IMCOT",
    "IMSEC",
    "IMCSC",
    "IMSINH",
    "IMCOSH",
    "IMSECH",
    "IMCSCH",
}
HUGE = {"1E200+1E200i", "123456789.123456789+1i"}  # Excel's sine and cosine lose their digits here
for fn in UNARY:
    skip = (
        HUGE | {"1E-20+1E-20i", "1e5i"}
        if fn in TRIG - {"IMSIN", "IMCOS", "IMEXP"}
        else HUGE
        if fn in TRIG
        else set()
    )
    CASES += [
        (f"={fn}({quoted(t)})", 1e-9) if fn in TRIG else f"={fn}({quoted(t)})"
        for t in COMPLEX_TEXTS
        if t not in skip
    ]
    CASES += [f"={fn}({c})" for c, _ in INPUT_CELLS]
    CASES += [
        f"={fn}({n})"
        for n in ("5", "-2.5", "0", "TRUE", "1E-300", "0.5", *(() if fn in TRIG else ("1E300",)))
    ]
BINARY_TEXTS = [
    "3+4i",
    "1-2i",
    "5",
    "i",
    "0",
    "-2.5+0.5j",
    "3+4j",
    "1E-5+1E5i",
    "0+0i",
    "7i",
    "abc",
]
for fn in ("IMSUB", "IMDIV", "IMSUM", "IMPRODUCT"):
    CASES += [f"={fn}({quoted(a)},{quoted(b)})" for a in BINARY_TEXTS for b in BINARY_TEXTS]
CASES += [
    f"={fn}({quoted(a)},{c})"
    for fn in ("IMSUB", "IMDIV", "IMSUM", "IMPRODUCT")
    for a in ("1+i",)
    for c, _ in INPUT_CELLS
]
CASES += [
    f"={fn}({quoted(a)},{quoted(b)},{quoted(c)})"
    for fn in ("IMSUM", "IMPRODUCT")
    for a, b, c in (
        ("1+i", "2-3i", "4"),
        ("1+i", "i", "j"),
        ("i", "i", "i"),
        ("2", "3", "4i"),
        ("1+i", "", "2"),
    )
]
CASES += (
    [f"={fn}(Data!Q2:Q6)" for fn in ("IMSUM", "IMPRODUCT")]
    + [f"={fn}(Data!C2:C6,{quoted('i')})" for fn in ("IMSUM", "IMPRODUCT")]
    + [f"={fn}(Data!F2:F5)" for fn in ("IMSUM", "IMPRODUCT")]
    + [f"={fn}(Data!G2:G3)" for fn in ("IMSUM", "IMPRODUCT")]
)
CASES += (
    [f'={fn}({{"1+i","2";"3i","4"}})' for fn in ("IMSUM", "IMPRODUCT")]
    + [f"={fn}({{1,2,3}})" for fn in ("IMSUM", "IMPRODUCT")]
    + [f"={fn}()" for fn in ("IMSUM", "IMPRODUCT")]
    if False
    else []
)
CASES += [f'={fn}({{"1+i","2";"3i","4"}})' for fn in ("IMSUM", "IMPRODUCT")] + [
    f"={fn}({{1,2,3}})" for fn in ("IMSUM", "IMPRODUCT")
]
for a in ("3+4i", "1+i", "i", "-1", "2", "0", "-8", "1-2j", "0.5", "1E-5i", "abc", "1E10+1E10i"):
    CASES += [
        f"=IMPOWER({quoted(a)},{p})"
        for p in (
            "2",
            "3",
            "0",
            "-1",
            "0.5",
            "1.5",
            "-2.5",
            "10",
            "100",
            "0.1",
            "TRUE",
            '"2"',
            '"a"',
            "1E-5",
            "40",
            "-40",
        )
    ]
CASES += [
    "=IMPOWER(Data!G2,2)",
    '=IMPOWER("1+i",Data!G2)',
    "=IMPOWER(Data!Q3,2)",
    '=IMPOWER("1+i",Data!Q3)',
    "=IMPOWER(2,2)",
    '=IMPOWER(2,"1+i")',
]
for r, i in (
    (3, 4),
    (0, 0),
    (-1, 0),
    (0, -1),
    (1, 1),
    (1.5, -2.5),
    (1e-5, 1e5),
    (1e10, 1e-10),
    (123456789.123456789, 0.000123456789),
    (0.1, 0.2),
    (-0.0, 3),
):
    CASES += [
        f"=COMPLEX({r},{i})",
        f'=COMPLEX({r},{i},"j")',
        f'=COMPLEX({r},{i},"i")',
        f'=COMPLEX({r},{i},"")',
        f'=COMPLEX({r},{i},"k")',
        f'=COMPLEX({r},{i},"I")',
        f'=COMPLEX({r},{i},"J")',
        f"=COMPLEX({r},{i},1)",
    ]
CASES += [
    '=COMPLEX("3","4")',
    "=COMPLEX(TRUE,1)",
    '=COMPLEX("a",1)',
    "=COMPLEX(Data!Q3,1)",
    "=COMPLEX(1,Data!Q3)",
    "=COMPLEX(Data!G2,1)",
    "=COMPLEX(1,2,Data!Q3)",
    "=COMPLEX(1,2,Data!A2)",
    '=COMPLEX(1,0,"j")',
    '=COMPLEX(0,0,"j")',
    '=COMPLEX(0,1,"j")',
    "=COMPLEX(0,-1)",
    "=COMPLEX(1E100,1E100)",
    "=COMPLEX(1E-100,1)",
    "=COMPLEX(-1E-300,2)",
    "=COMPLEX(1/3,2/3)",
]
CASES += [
    '=IMREAL("3+4i")+1',
    '=IMABS(IMSUM("3+4i","1"))',
    '=IMREAL(IMPRODUCT("1+i","1+i"))',
    '=IMSUM("3+4i","1+2i")&""',
    '=LEN(IMEXP("1+i"))',
    '=IMAGINARY(IMSQRT("-4"))',
    '=IMSQRT("-4")',
    '=IMSQRT("-1")',
    "=IMSQRT(-2)",
    "=IMEXP(0)",
    "=IMLN(1)",
    "=IMLN(-1)",
    "=IMSIN(0)",
    "=IMCOS(0)",
    "=IMCOSH(0)",
    '=IMPOWER("i",2)',
    '=IMPOWER("i",4)',
    '=IMDIV("1","i")',
    '=IMDIV("i","i")',
    '=IMPRODUCT("i","i")',
    '=IMPRODUCT("i","-i")',
    "=IMLOG10(100)",
    "=IMLOG2(8)",
    '=IMLOG10("i")',
    "=IMLOG2(-8)",
]

# ---- CONVERT
GROUPS = {
    "mass": [
        "g",
        "sg",
        "lbm",
        "u",
        "ozm",
        "grain",
        "cwt",
        "shweight",
        "uk_cwt",
        "lcwt",
        "hweight",
        "stone",
        "ton",
        "uk_ton",
        "LTON",
        "brton",
    ],
    "distance": [
        "m",
        "mi",
        "Nmi",
        "in",
        "ft",
        "yd",
        "ang",
        "ell",
        "ly",
        "parsec",
        "pc",
        "Pica",
        "pica",
        "Picapt",
        "survey_mi",
    ],
    "time": ["yr", "day", "d", "hr", "mn", "min", "sec", "s"],
    "pressure": ["Pa", "p", "atm", "at", "mmHg", "psi", "Torr"],
    "force": ["N", "dyn", "dy", "lbf", "pond"],
    "energy": ["J", "e", "c", "cal", "eV", "ev", "HPh", "hh", "Wh", "wh", "flb", "BTU", "btu"],
    "power": ["HP", "h", "PS", "W", "w"],
    "magnetism": ["T", "ga"],
    "temperature": ["C", "cel", "F", "fah", "K", "kel", "Rank", "Reau"],
    "volume": [
        "tsp",
        "tspm",
        "tbs",
        "oz",
        "cup",
        "pt",
        "us_pt",
        "uk_pt",
        "qt",
        "uk_qt",
        "gal",
        "uk_gal",
        "l",
        "L",
        "lt",
        "ang3",
        "ang^3",
        "ft3",
        "ft^3",
        "in3",
        "in^3",
        "ly3",
        "ly^3",
        "m3",
        "m^3",
        "mi3",
        "mi^3",
        "yd3",
        "yd^3",
        "Nmi3",
        "Nmi^3",
        "Pica3",
        "Picapt3",
        "Pica^3",
        "Picapt^3",
        "GRT",
        "regton",
        "MTON",
    ],
    "area": [
        "uk_acre",
        "us_acre",
        "ang2",
        "ang^2",
        "ar",
        "ft2",
        "ft^2",
        "ha",
        "in2",
        "in^2",
        "ly2",
        "ly^2",
        "m2",
        "m^2",
        "Morgen",
        "mi2",
        "mi^2",
        "Nmi2",
        "Nmi^2",
        "Pica2",
        "Picapt2",
        "Pica^2",
        "Picapt^2",
        "yd2",
        "yd^2",
    ],
    "information": ["bit", "byte"],
    "speed": ["admkn", "kn", "m/h", "m/hr", "m/s", "m/sec", "mph"],
}
PREFIXABLE = [
    "g",
    "u",
    "m",
    "ang",
    "ly",
    "parsec",
    "pc",
    "sec",
    "s",
    "Pa",
    "p",
    "atm",
    "at",
    "mmHg",
    "N",
    "dyn",
    "dy",
    "pond",
    "J",
    "e",
    "c",
    "cal",
    "eV",
    "ev",
    "Wh",
    "wh",
    "W",
    "w",
    "T",
    "ga",
    "uk_pt",
    "l",
    "L",
    "lt",
    "ang3",
    "ang^3",
    "m3",
    "m^3",
    "ang2",
    "ang^2",
    "ar",
    "m2",
    "m^2",
    "bit",
    "byte",
    "m/h",
    "m/hr",
    "m/s",
    "m/sec",
    "mph",
    "K",
    "kel",
]
PREFIXES = [
    "Y",
    "Z",
    "E",
    "P",
    "T",
    "G",
    "M",
    "k",
    "h",
    "da",
    "e",
    "d",
    "c",
    "m",
    "u",
    "n",
    "p",
    "f",
    "a",
    "z",
    "y",
    "ki",
    "Mi",
    "Gi",
    "Ti",
    "Pi",
    "Ei",
    "Zi",
    "Yi",
]
for members in GROUPS.values():
    for index, unit in enumerate(members):
        for other in {
            members[0],
            members[index - 1],
            members[(index + 1) % len(members)],
            members[-1],
        }:
            CASES.append(f'=CONVERT(7.25,"{unit}","{other}")')
        CASES.append(f'=CONVERT(-3,"{unit}","{members[0]}")')
WIDE = {
    "m",
    "g",
    "bit",
    "byte",
    "m2",
    "m3",
    "l",
    "K",
    "s",
    "Pa",
    "W",
    "N",
}  # every prefix; others a few
for unit in PREFIXABLE:
    base = next(m[0] for m in GROUPS.values() if unit in m)
    for prefix in PREFIXES if unit in WIDE else ("k", "m", "da", "c", "u", "G"):
        CASES.append(f'=CONVERT(1,"{prefix}{unit}","{unit}")')
        CASES.append(f'=CONVERT(3.5,"{unit}","{prefix}{unit}")')
    for prefix in ("k", "m"):
        CASES.append(f'=CONVERT(2,"{prefix}{unit}","{base}")')
temperatures = GROUPS["temperature"]
for value in (-300, -40, 0, 32, 37, 100, 212, 1e6):
    for a in temperatures:
        for b in temperatures:
            CASES.append(f'=CONVERT({value},"{a}","{b}")')
for a in ("kK", "mK", "kC", "mC", "kF", "kRank", "kReau", "mkel", "kcel"):
    CASES += [f'=CONVERT(300,"{a}","K")', f'=CONVERT(300,"K","{a}")', f'=CONVERT(300,"{a}","C")']
CASES += [
    '=CONVERT(1,"m","ft")',
    '=CONVERT(1,"ft","m")',
    '=CONVERT(1,"kg","lbm")',
    '=CONVERT(1,"lbm","kg")',
    '=CONVERT(1,"km","mi")',
    '=CONVERT(1,"mph","kn")',
    '=CONVERT(1,"mm","in")',
    '=CONVERT(1,"m","kg")',
    '=CONVERT(1,"xx","m")',
    '=CONVERT(1,"m","xx")',
    '=CONVERT(1,"M","m")',
    '=CONVERT(1,"Km","m")',
    '=CONVERT(1,"KM","m")',
    '=CONVERT(1,"MB","byte")',
    '=CONVERT(1,"Mibyte","byte")',
    '=CONVERT(1,"kibyte","bit")',
    '=CONVERT(1,"kibit","bit")',
    '=CONVERT(1,"kift","ft")',
    '=CONVERT(1,"kft","ft")',
    '=CONVERT(1,"kibyte","kbyte")',
    '=CONVERT(1,"dam","m")',
    '=CONVERT(1,"dm","m")',
    '=CONVERT(1,"hm","m")',
    '=CONVERT(1,"mm2","m2")',
    '=CONVERT(1,"km2","m2")',
    '=CONVERT(1,"km^2","m^2")',
    '=CONVERT(1,"km3","m3")',
    '=CONVERT(1,"km3","l")',
    '=CONVERT(1,"kl","m3")',
    '=CONVERT(1,"ml","cm3")',
    '=CONVERT(1,"cm3","ml")',
    '=CONVERT(1,"m3","l")',
    '=CONVERT(1,"kmin","min")',
    '=CONVERT(1,"ms","s")',
    '=CONVERT(1,"ks","s")',
    '=CONVERT(1,"hr","s")',
    '=CONVERT(1,"hs","s")',
    '=CONVERT(1,"hPa","Pa")',
    '=CONVERT(1,"kPa","psi")',
    '=CONVERT(1,"mmHg","Pa")',
    '=CONVERT(1,"kmmHg","Pa")',
    '=CONVERT(1,"kcal","J")',
    '=CONVERT(1,"kc","J")',
    '=CONVERT(1,"kWh","J")',
    '=CONVERT(1,"MW","HP")',
    '=CONVERT(1,"kW","W")',
    '=CONVERT(1,"mT","ga")',
    '=CONVERT(1,"kyd","m")',
    '=CONVERT(1,"kmi","m")',
    '=CONVERT(1,"kyr","s")',
    '=CONVERT(1,"kpt","l")',
    '=CONVERT(1,"kuk_pt","l")',
    '=CONVERT(1,"kuk_qt","l")',
    '=CONVERT(1,"kstone","g")',
    '=CONVERT(1,"kg","stone")',
    '=CONVERT(1,"m/s","km/h")',
    '=CONVERT(1,"km/h","m/s")',
    '=CONVERT(1,"kmph","m/s")',
    '=CONVERT(1,"","m")',
    '=CONVERT(1,"m","")',
    '=CONVERT(1,"m"," m")',
    '=CONVERT(1,"m","m ")',
    '=CONVERT(1,m,"m")',
    '=CONVERT(1,1,"m")',
    '=CONVERT(1,"m",1)',
    '=CONVERT(1,TRUE,"m")',
    '=CONVERT(1,Data!Q3,"m")',
    '=CONVERT(1,"m",Data!Q3)',
    '=CONVERT(1,Data!G2,"m")',
    '=CONVERT(1,"m",Data!G2)',
    '=CONVERT(Data!G2,"m","ft")',
    '=CONVERT("2","m","ft")',
    '=CONVERT("a","m","ft")',
    '=CONVERT(TRUE,"m","ft")',
    '=CONVERT(Data!Q3,"m","ft")',
    '=CONVERT(0,"m","ft")',
    '=CONVERT(-1,"m","ft")',
    '=CONVERT(1E300,"km","m")',
    '=CONVERT(1E300,"Ym","ym")',
    '=CONVERT(1E-300,"ym","Ym")',
    '=CONVERT(1,"m","m")',
    '=CONVERT(1,"Pica","pica")',
    '=CONVERT(1,"Pica","Picapt")',
    '=CONVERT(1,"cup","tbs")',
    '=CONVERT({1,2},"m","ft")',
    '=CONVERT(1,"ang","m")',
    '=CONVERT(1,"ly","m")',
    '=CONVERT(1,"pc","ly")',
    '=CONVERT(1,"parsec","ly")',
    '=CONVERT(1,"eV","J")',
    '=CONVERT(1,"ev","kWh")',
    '=CONVERT(1,"HPh","Wh")',
    '=CONVERT(1,"BTU","cal")',
    '=CONVERT(1,"c","cal")',
    '=CONVERT(1,"e","J")',
    '=CONVERT(1,"ee","J")',
    '=CONVERT(1,"de","J")',
    '=CONVERT(1,"daJ","J")',
    '=CONVERT(1,"dag","g")',
    '=CONVERT(1,"dal","l")',
    '=CONVERT(1,"das","s")',
    '=CONVERT(1,"daday","s")',
]
CASES += [
    '=CONVERT(1,"pc","c")',
    '=CONVERT(1,"c","pc")',
    '=CONVERT(1,"pc","cal")',
    '=CONVERT(1,"pc","m")',
    '=CONVERT(1,"pm","m")',
    '=CONVERT(1,"pm","pc")',
    '=CONVERT(1,"Pa","ac")',
    '=CONVERT(1,"mmm","m")',
    '=CONVERT(1,"hh","h")',
    '=CONVERT(1,"hh","W")',
    '=CONVERT(1,"dy","dyn")',
    '=CONVERT(1,"ds","d")',
    '=CONVERT(1,"dd","s")',
    '=CONVERT(1,"mn","mm")',
]

# ---- Bessel functions: Excel is accurate to about 1e-8, for J near x = 10 to 5e-5, and for large x of I with n > 1 to 1 part in 10
BESSEL_X = [
    "0",
    "1E-10",
    "0.001",
    "0.1",
    "0.5",
    "1",
    "1.5",
    "2",
    "2.404825557695773",
    "3.7",
    "5",
    "7.5",
    "10",
    "12.3",
    "20",
    "30.5",
    "45",
    "59.9",
    "60.1",
    "75",
    "100",
    "150",
    "250",
    "700",
    "1000",
    "-0.5",
    "-1",
    "-3.7",
    "-10",
    '"2"',
    "TRUE",
    "Data!Q3",
    "Data!G2",
]
BESSEL_N = ["0", "1", "2", "5", "20", "1.9", "-1", "TRUE", "Data!Q3"]
BESSEL_TOLERANCE = {"BESSELJ": 1e-4, "BESSELY": 1e-6, "BESSELI": 1e-6, "BESSELK": 1e-6}
for fn in ("BESSELJ", "BESSELY", "BESSELI", "BESSELK"):
    for x in BESSEL_X:
        for n in BESSEL_N:
            if (
                fn == "BESSELI"
                and x in ("150", "250", "700", "1000")
                and n not in ("0", "1", "-1", "TRUE", "Data!Q3")
            ):
                continue
            slack = (
                2e-3
                if fn == "BESSELI"
                and x in ("45", "59.9", "60.1", "75", "100", "150", "250", "700", "1000")
                else BESSEL_TOLERANCE[fn]
            )
            CASES.append((f"={fn}({x},{n})", slack))
CASES += [
    (f"={fn}({x},0)", 1e-4)
    for fn in ("BESSELJ", "BESSELY")
    for x in (
        "1E5",
        "1E6",
        "1E7",
        "1E8",
        "134217727",
        "134217728",
        "1E10",
        "1E15",
        "-1E5",
        "-134217728",
    )
]
CASES += [
    (f"=BESSELI({x},0)", 1e-6) for x in ("709", "709.7", "709.71", "710", "-710", "-709.7")
] + [(f"=BESSELK({x},0)", 1e-6) for x in ("1E-300", "700", "745", "746", "1000", "1E5")]
CASES += [
    (f"={fn}({x},{n})", 1e-6)
    for fn in ("BESSELJ", "BESSELI", "BESSELK", "BESSELY")
    for x, n in (("5", "200"), ("5", "1000"), ("1E-10", "3"), ("1E-5", "50"), ("1E-300", "0"))
]

# ---- ENCODEURL
CASES += [
    '=ENCODEURL("a b")',
    '=ENCODEURL("abcXYZ019-_.~")',
    '=ENCODEURL("!*\'();:@&=+$,/?#[]")',
    '=ENCODEURL("%")',
    "=ENCODEURL(UNICHAR(233)&UNICHAR(8364)&UNICHAR(128512))",
    '=ENCODEURL("")',
    "=ENCODEURL(123.5)",
    "=ENCODEURL(TRUE)",
    "=ENCODEURL(Data!Q3)",
    "=ENCODEURL(Data!G2)",
    '=ENCODEURL("a"&CHAR(10)&"b")',
    r'=ENCODEURL("<>""{}|\^`")',
    '=ENCODEURL("https://x.com/a b?c=d&e="&UNICHAR(233))',
    "=ENCODEURL(1E-10)",
    "=ENCODEURL(1E+20)",
    '=ENCODEURL(" ")',
    '=ENCODEURL("a+b")',
    "=ENCODEURL(UNICHAR(0))" if False else '=ENCODEURL("tab"&CHAR(9))',
    '=ENCODEURL("\u00e9")',
    "=ENCODEURL(Data!A2)",
    "=ENCODEURL(Data!R3)",
    "=ENCODEURL(Data!R4)",
]
