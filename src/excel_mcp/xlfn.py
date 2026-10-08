"""Storage names of functions added after Excel 2007.

Excel stores these with a prefix (``_xlfn.IFS``). A formula written without it is
reported as #NAME? when the file is opened, and some (``SORT``) stop the file from
opening at all, so formulas are prefixed when they are written. The names that LET
declares are stored with a ``_xlpm.`` prefix.
"""

import re
from dataclasses import dataclass, field

from openpyxl.formula import Tokenizer
from openpyxl.formula.tokenizer import Token

_XLFN = [
    "ACOT",
    "ACOTH",
    "AGGREGATE",
    "ARABIC",
    "BASE",
    "BETA.DIST",
    "BETA.INV",
    "BINOM.DIST",
    "BINOM.DIST.RANGE",
    "BINOM.INV",
    "BITAND",
    "BITLSHIFT",
    "BITOR",
    "BITRSHIFT",
    "BITXOR",
    "CEILING.MATH",
    "CEILING.PRECISE",
    "CHISQ.DIST",
    "CHISQ.DIST.RT",
    "CHISQ.INV",
    "CHISQ.INV.RT",
    "CHISQ.TEST",
    "CHOOSECOLS",
    "CHOOSEROWS",
    "COMBINA",
    "CONCAT",
    "CONFIDENCE.NORM",
    "CONFIDENCE.T",
    "COT",
    "COTH",
    "COVARIANCE.P",
    "COVARIANCE.S",
    "CSC",
    "CSCH",
    "DAYS",
    "DECIMAL",
    "DROP",
    "ERF.PRECISE",
    "ERFC.PRECISE",
    "EXPAND",
    "EXPON.DIST",
    "F.DIST",
    "F.DIST.RT",
    "F.INV",
    "F.INV.RT",
    "F.TEST",
    "FLOOR.MATH",
    "FLOOR.PRECISE",
    "FORECAST.ETS",
    "FORECAST.LINEAR",
    "FORMULATEXT",
    "GAMMA",
    "GAMMA.DIST",
    "GAMMA.INV",
    "GAMMALN.PRECISE",
    "GAUSS",
    "HSTACK",
    "HYPGEOM.DIST",
    "IFNA",
    "IFS",
    "ISFORMULA",
    "ISOWEEKNUM",
    "LAMBDA",
    "LET",
    "LOGNORM.DIST",
    "LOGNORM.INV",
    "MAKEARRAY",
    "MAP",
    "MAXIFS",
    "MINIFS",
    "MODE.MULT",
    "MODE.SNGL",
    "MUNIT",
    "NEGBINOM.DIST",
    "NORM.DIST",
    "NORM.INV",
    "NORM.S.DIST",
    "NORM.S.INV",
    "NUMBERVALUE",
    "PDURATION",
    "PERCENTILE.EXC",
    "PERCENTILE.INC",
    "PERCENTRANK.EXC",
    "PERCENTRANK.INC",
    "PERMUTATIONA",
    "PHI",
    "POISSON.DIST",
    "QUARTILE.EXC",
    "QUARTILE.INC",
    "RANDARRAY",
    "RANK.AVG",
    "RANK.EQ",
    "REDUCE",
    "RRI",
    "SCAN",
    "SEC",
    "SECH",
    "SEQUENCE",
    "SHEET",
    "SHEETS",
    "SKEW.P",
    "SORTBY",
    "STDEV.P",
    "STDEV.S",
    "SWITCH",
    "T.DIST",
    "T.DIST.2T",
    "T.DIST.RT",
    "T.INV",
    "T.INV.2T",
    "T.TEST",
    "TAKE",
    "TEXTAFTER",
    "TEXTBEFORE",
    "TEXTJOIN",
    "TEXTSPLIT",
    "TOCOL",
    "TOROW",
    "UNICHAR",
    "UNICODE",
    "UNIQUE",
    "VAR.P",
    "VAR.S",
    "VSTACK",
    "WEIBULL.DIST",
    "WRAPCOLS",
    "WRAPROWS",
    "XLOOKUP",
    "XMATCH",
    "XOR",
    "Z.TEST",
]
_XLWS = ["FILTER", "SORT"]

FUTURE_FUNCTIONS: dict[str, str] = {
    **dict.fromkeys(_XLFN, "_xlfn."),
    **dict.fromkeys(_XLWS, "_xlfn._xlws."),
}


_PREFIXES = re.compile(r"^(?:_xl(?:fn|ws|pm)\.)+")


@dataclass
class _Call:
    name: str
    start: int
    args: list[list[int]] = field(default_factory=lambda: [[]])


def add_prefixes(formula: str) -> str:
    """Return the formula with the storage prefixes Excel expects on functions and LET names."""
    tokenizer = Tokenizer(formula)
    tokens = tokenizer.items
    changed = False
    names: set[int] = set()
    stack: list[_Call] = []
    for index, token in enumerate(tokens):
        if token.type == Token.FUNC and token.subtype == Token.OPEN:
            name = token.value.removesuffix("(").upper()
            if prefix := FUTURE_FUNCTIONS.get(name):
                token.value = prefix + token.value
                changed = True
            if stack:
                stack[-1].args[-1].append(index)
            stack.append(_Call(name, index))
        elif token.type == Token.FUNC and token.subtype == Token.CLOSE and stack:
            call = stack.pop()
            if call.name == "LET":
                declared = _declared_names(tokens, call)
                names.update(
                    position
                    for position in range(call.start, index)
                    if _is_name(tokens[position]) and tokens[position].value.casefold() in declared
                )
        elif token.type == Token.SEP and token.subtype == Token.ARG and stack:
            stack[-1].args.append([])
        elif stack:
            stack[-1].args[-1].append(index)
    for position in names:
        tokens[position].value = "_xlpm." + tokens[position].value
        changed = True
    return tokenizer.render() if changed else formula


def remove_prefixes(formula: str) -> str:
    """Return the formula as Excel shows it, without the prefixes `add_prefixes` adds."""
    tokenizer = Tokenizer(formula)
    for token in tokenizer.items:
        if token.type == Token.FUNC or token.subtype == Token.RANGE:
            token.value = _PREFIXES.sub("", token.value)
    return tokenizer.render()


def _is_name(token: Token) -> bool:
    return (
        token.type == Token.OPERAND
        and token.subtype == Token.RANGE
        and not token.value.startswith("_xlpm.")
    )


def _declared_names(tokens: list[Token], call: _Call) -> set[str]:
    """The names a LET call declares: its odd-numbered arguments, except the last."""
    declared = set()
    for number, indexes in enumerate(call.args[:-1]):
        parts = [i for i in indexes if tokens[i].type != Token.WSPACE]
        if number % 2 == 0 and len(parts) == 1 and _is_name(tokens[parts[0]]):
            declared.add(tokens[parts[0]].value.casefold())
    return declared
