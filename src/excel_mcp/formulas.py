"""Formula safety policy.

Every formula the server writes (cell values, conditional formats, data
validation rules) passes through `check_formula`. Formulas are tokenized with
openpyxl's Excel tokenizer rather than matched with regular expressions, and
anything that cannot be tokenized is rejected.

Blocked are functions that reach the network or other programs, leak
information about the host, or build references at runtime, plus references to
external workbooks and DDE links.
"""

from openpyxl.formula import Tokenizer
from openpyxl.formula.tokenizer import Token, TokenizerError

from excel_mcp.errors import UnsafeFormulaError

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
    name to MAP or BYROW) and ``Sheet1!WEBSERVICE`` all become ``WEBSERVICE``.
    """
    name = token_value.removesuffix("(").strip().upper().lstrip("@")
    name = name.rpartition("!")[2]
    while name.startswith(_FUNCTION_PREFIXES):
        name = name.split(".", 1)[1]
    return name


def check_formula(formula: str) -> None:
    """Raise `UnsafeFormulaError` unless ``formula`` (starting with ``=``) is allowed."""
    if not formula.startswith("="):
        raise UnsafeFormulaError(f"Formula must start with '=': {formula!r}.")
    try:
        tokens = Tokenizer(formula).items
    except TokenizerError as error:
        raise UnsafeFormulaError(f"Formula could not be parsed: {error}.") from None

    for token in tokens:
        _check_token(token)


def _check_token(token: Token) -> None:
    if "|" in token.value and token.subtype != Token.TEXT:
        raise UnsafeFormulaError("DDE links ('|') are not allowed in formulas.")
    # A blocked name is rejected even without its "(": the tokenizer reads
    # "=WEBSERVICE (A1)" as a name followed by a parenthesis.
    is_function = token.type == Token.FUNC and token.subtype == Token.OPEN
    if is_function or token.subtype == Token.RANGE:
        name = normalize_function_name(token.value)
        if name in BLOCKED_FUNCTIONS:
            raise UnsafeFormulaError(
                f"Function {name} is not allowed because it can access the network, "
                "other programs or host information."
            )
    if token.subtype == Token.RANGE and _is_external_reference(token.value):
        raise UnsafeFormulaError(f"References to other workbooks are not allowed: {token.value!r}.")


def _is_external_reference(reference: str) -> bool:
    # External references put the workbook in brackets before the sheet, as in
    # "[1]Sheet1!A1" or "'C:\dir\[book.xlsx]Sheet1'!A1". Table references such
    # as "Table1[Sales]" also use brackets but have no sheet part.
    sheet_part, separator, _ = reference.rpartition("!")
    return bool(separator) and "[" in sheet_part
