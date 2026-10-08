"""Formula safety policy.

Every formula the server writes (cell values, conditional formats, data
validation rules) passes through `check_formula`. Formulas are tokenized with
openpyxl's Excel tokenizer rather than matched with regular expressions, and
anything that cannot be tokenized is rejected.

Blocked are functions that reach the network or other programs, leak
information about the host, or build references at runtime, plus DDE links.
A reference may only name sheets of the workbook it is written to: anything
else before a "!" (a file, a path, a URL, "[1]") points at another workbook.
"""

from collections.abc import Iterable, Iterator

from openpyxl.formula import Tokenizer
from openpyxl.formula.tokenizer import Token, TokenizerError

from excel_mcp.errors import UnsafeFormulaError
from excel_mcp.xlfn import add_prefixes

BLOCKED_FUNCTIONS = frozenset(
    {
        # Network access
        "WEBSERVICE",
        "FILTERXML",
        "IMAGE",
        "HYPERLINK",
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
        # Excel 4.0 macro functions
        "EXEC",
        "EXECUTE",
        "EVALUATE",
        "FOPEN",
        "FWRITE",
        "FWRITELN",
        "FREAD",
        "FREADLN",
        "FCLOSE",
        "GET.WORKSPACE",
        "GET.DOCUMENT",
        "GET.CELL",
        # Host information and runtime references
        "INFO",
        "CELL",
        "INDIRECT",
    }
)

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
    """Raise `UnsafeFormulaError` unless ``formula`` (starting with ``=``) is allowed.

    ``sheet_names`` are the sheets of the workbook the formula is written to.
    """
    if not formula.startswith("="):
        raise UnsafeFormulaError(f"Formula must start with '=': {formula!r}.")
    try:
        tokens = Tokenizer(formula).items
    except TokenizerError as error:
        raise UnsafeFormulaError(f"Formula could not be parsed: {error}.") from None

    sheets = {name.casefold(): name for name in sheet_names}
    for token in tokens:
        _check_token(token, sheets)


def storable_formula(formula: str, sheet_names: Iterable[str]) -> str:
    """Check a formula and return it as it is stored in a file (see `add_prefixes`).

    Every formula that is written anywhere in a workbook goes through here.
    """
    check_formula(formula, sheet_names)
    return add_prefixes(formula)


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
            f"Function {name} is not allowed because it can access the network, "
            "other programs or host information."
        )
    for qualifier in _qualifiers(token.value):
        # Sheet names are case-insensitive, and a 3D reference spans "First:Last".
        for sheet in qualifier.split(":"):
            if sheet.casefold() not in sheets:
                raise UnsafeFormulaError(
                    f"{sheet!r} in {token.value!r} is not a sheet of this workbook, and "
                    "references to other workbooks are not allowed. "
                    f"Sheets: {', '.join(sheets.values())}."
                )


def _qualifiers(reference: str) -> Iterator[str]:
    """The unquoted names before each "!" in a reference, e.g. "My Sheet" in "'My Sheet'!A1"."""
    parts = _split_top_level(reference, ":")
    for previous, part in zip([None, *parts], parts, strict=False):
        qualifier = _split_top_level(part, "!")
        if len(qualifier) == 1:
            continue
        yield _unquote(qualifier[0])
        # An unquoted 3D reference such as "Sheet1:Sheet3!A1" starts in the part before.
        if previous is not None and len(_split_top_level(previous, "!")) == 1:
            yield _unquote(previous)


def _split_top_level(text: str, separator: str) -> list[str]:
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


def _unquote(name: str) -> str:
    if len(name) > 1 and name.startswith("'") and name.endswith("'"):
        return name[1:-1].replace("''", "'")
    return name
