"""Formula safety policy.

Every formula the server writes (cell values, conditional formats, data
validation rules) passes through `check_formula`. Formulas are tokenized with
openpyxl's Excel tokenizer rather than matched with regular expressions, and
anything that cannot be tokenized is rejected.

Blocked are functions that reach the network or other programs, Excel 4.0
macro functions (`excel_mcp.xlm`) and DDE links. HYPERLINK is allowed with a
literal link only. INDIRECT, CELL and INFO are allowed: they only read the
open workbook and the host, and without a function that sends data out there
is nothing to leak it with.
A reference may only name sheets of the workbook it is written to: anything
else before a "!" (a file, a path, a URL, "[1]") points at another workbook.
"""

import re
from collections.abc import Iterable, Iterator

from openpyxl.formula import Tokenizer
from openpyxl.formula.tokenizer import Token, TokenizerError

from excel_mcp.errors import InvalidFormulaError, UnsafeFormulaError
from excel_mcp.formula_syntax import check_syntax, invalid_formula
from excel_mcp.spill import store_spills
from excel_mcp.text import quoted
from excel_mcp.xlfn import add_prefixes
from excel_mcp.xlm import is_macro_function

BLOCKED_FUNCTIONS = frozenset(
    {
        # Network access
        "WEBSERVICE",
        "FILTERXML",
        "IMAGE",
        "STOCKHISTORY",
        "TRANSLATE",
        "DETECTLANGUAGE",
        "COPILOT",
        "PY",
        # Google Sheets network functions, for files opened there
        "IMPORTDATA",
        "IMPORTFEED",
        "IMPORTHTML",
        "IMPORTRANGE",
        "IMPORTXML",
        # External programs and data connections
        "RTD",
        "DDE",
        "CALL",
        "REGISTER",
        "REGISTER.ID",
        "SQL.REQUEST",
        "CUBEKPIMEMBER",
        "CUBEMEMBER",
        "CUBEMEMBERPROPERTY",
        "CUBERANKEDMEMBER",
        "CUBESET",
        "CUBESETCOUNT",
        "CUBEVALUE",
    }
)

# Brackets and slashes make a link point into another workbook or file.
_LINK_ESCAPES = re.compile(r"[\\[\]]|://")
_FUNCTION_PREFIXES = ("_XLFN.", "_XLWS.", "_XLUDF.", "_XLETA.")


def normalize_function_name(token_value: str) -> str:
    """Reduce a function or name token to the bare function name Excel would call.

    ``@_xlfn._xlfn.webservice(``, ``_xleta.WEBSERVICE`` (a function passed by
    name to MAP or BYROW), ``Sheet1!WEBSERVICE`` and ``A1:WEBSERVICE(`` all
    become ``WEBSERVICE``.
    """
    name = token_value.removesuffix("(").strip().upper()
    if token_value.endswith("("):
        # A call can follow a range operator. Range tokens keep their colon,
        # because PY, DDE and RTD are also column names, as in "A:PY".
        name = name.rpartition(":")[2]
    name = name.rpartition("!")[2].strip().lstrip("@")
    while name.startswith(_FUNCTION_PREFIXES):
        name = name.split(".", 1)[1]
    return name


def check_formula(formula: str, sheet_names: Iterable[str]) -> None:
    """Raise unless ``formula`` (starting with ``=``) is valid Excel syntax and allowed.

    Raises `InvalidFormulaError` for syntax and `UnsafeFormulaError` for the safety policy.
    ``sheet_names`` are the sheets of the workbook the formula is written to.
    """
    if not formula.startswith("="):
        raise InvalidFormulaError(f"Formula must start with '=': {formula!r}.")
    tokens = _tokenize(formula)
    sheets = {name.casefold(): name for name in sheet_names}
    for index, token in enumerate(tokens):
        _check_token(token, sheets)
        if (arguments := _call_arguments(tokens, index)) is not None:
            _check_call(normalize_function_name(token.value), tokens[arguments:], sheets)
    check_syntax(formula, tokens)


def _tokenize(formula: str) -> list[Token]:
    try:
        return Tokenizer(store_spills(formula)).items
    except TokenizerError as error:
        reason = str(error).removesuffix(f" in {formula!r}").removesuffix(f" in {formula}")
        if "parsing string" in reason:
            reason = "unterminated text, close it with a quote"
        raise invalid_formula(formula, reason) from None
    except IndexError:
        raise invalid_formula(formula, "unmatched ')'") from None


def storable_formula(formula: str, sheet_names: Iterable[str]) -> str:
    """Check a formula and return it as it is stored in a file (see `add_prefixes`).

    Every formula that is written anywhere in a workbook goes through here.
    """
    check_formula(formula, sheet_names)
    return add_prefixes(store_spills(formula))


def storable_operand(operand: str, sheet_names: Iterable[str]) -> str:
    """`storable_formula` for rule operands and names, which are stored without the "="."""
    return storable_formula(f"={operand.removeprefix('=')}", sheet_names).removeprefix("=")


def _check_token(token: Token, sheets: dict[str, str]) -> None:
    if "|" in token.value and token.subtype != Token.TEXT:
        raise UnsafeFormulaError("DDE links ('|') are not allowed in formulas.")
    # A blocked name is rejected even without its "(": the tokenizer reads
    # "=WEBSERVICE (A1)" as a name followed by a parenthesis.
    is_function = token.type == Token.FUNC and token.subtype == Token.OPEN
    if not (is_function or token.subtype == Token.RANGE):
        return
    name = normalize_function_name(token.value)
    if name in BLOCKED_FUNCTIONS:
        raise UnsafeFormulaError(
            f"Function {name} is not allowed because it can access the network or other programs."
        )
    _check_qualifiers(token.value, sheets)


def _check_qualifiers(reference: str, sheets: dict[str, str]) -> None:
    for qualifier in _qualifiers(reference):
        # Sheet names are case-insensitive, and a 3D reference spans "First:Last".
        for sheet in qualifier.split(":"):
            if sheet.casefold() not in sheets:
                raise UnsafeFormulaError(
                    f"{sheet!r} in {reference!r} is not a sheet of this workbook, and "
                    "references to other workbooks are not allowed. "
                    f"Sheets: {quoted(sheets.values())}."
                )


def _call_arguments(tokens: list[Token], index: int) -> int | None:
    """Where the arguments of a call start, if the token at ``index`` names the function called.

    Names are only checked as calls when they are followed by "(" (the tokenizer reads
    "=FILES (A1)" as a name and a parenthesis) or passed by name to MAP and its relatives.
    """
    token = tokens[index]
    if token.type == Token.FUNC and token.subtype == Token.OPEN:
        return index + 1
    if token.subtype != Token.RANGE:
        return None
    following = next(
        (i for i in range(index + 1, len(tokens)) if tokens[i].type != Token.WSPACE), None
    )
    if following is not None and tokens[following].type == Token.PAREN:
        return following + 1 if tokens[following].subtype == Token.OPEN else None
    return index + 1 if token.value.upper().startswith("_XLETA.") else None


def _check_call(name: str, arguments: list[Token], sheets: dict[str, str]) -> None:
    if is_macro_function(name):
        raise UnsafeFormulaError(
            f"Function {name} is an Excel 4.0 macro function, which can read files, other "
            "workbooks or the application, and is not allowed."
        )
    if name == "HYPERLINK":
        _check_hyperlink(arguments, sheets)


def _check_hyperlink(arguments: list[Token], sheets: dict[str, str]) -> None:
    """Only a literal link can be clicked safely: one built from cell values could leak them."""
    first = [token for token in arguments if token.type != Token.WSPACE][:2]
    literal = len(first) == 2 and first[0].subtype == Token.TEXT
    if literal and first[1].subtype in (Token.ARG, Token.CLOSE):
        target = first[0].value[1:-1].replace('""', '"')
        if target.lower().startswith(("http://", "https://", "mailto:")):
            return
        if target.startswith("#") and not _LINK_ESCAPES.search(target):
            _check_qualifiers(target[1:], sheets)
            return
    raise UnsafeFormulaError(
        'HYPERLINK is not allowed unless its link is a literal text starting with "http://", '
        '"https://", "mailto:" or "#" (a place in this workbook, e.g. "#Sheet2!A1"). '
        "A link built from cells could send their contents to another server."
    )


def _qualifiers(reference: str) -> Iterator[str]:
    """The unquoted names before each "!" in a reference, e.g. "My Sheet" in "'My Sheet'!A1"."""
    parts = split_top_level(reference, ":")
    for previous, part in zip([None, *parts], parts, strict=False):
        qualifier = split_top_level(part, "!")
        if len(qualifier) == 1:
            continue
        yield unquote(qualifier[0])
        # An unquoted 3D reference such as "Sheet1:Sheet3!A1" starts in the part before.
        if previous is not None and len(split_top_level(previous, "!")) == 1:
            yield unquote(previous)


def split_top_level(text: str, separator: str) -> list[str]:
    """Split at ``separator`` outside quoted sheet names and outside brackets.

    Brackets hold table columns ("Table1[[Col1]:[Col2]]") or a workbook ("[1]"),
    and inside them "'" escapes the next character.
    """
    parts = []
    start = depth = 0
    quoted = escaped = False
    for index, character in enumerate(text):
        if escaped:
            escaped = False
        elif quoted:
            quoted = character != "'"
        elif depth:
            escaped = character == "'"
            depth += {"[": 1, "]": -1}.get(character, 0)
        elif character == "'":
            quoted = True
        elif character == "[":
            depth = 1
        elif character == separator:
            parts.append(text[start:index])
            start = index + 1
    parts.append(text[start:])
    return parts


def unquote(name: str) -> str:
    if len(name) > 1 and name.startswith("'") and name.endswith("'"):
        return name[1:-1].replace("''", "'")
    return name
