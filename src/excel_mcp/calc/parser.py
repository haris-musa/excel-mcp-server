"""Formula text to syntax tree, using openpyxl's tokenizer."""

import re
from collections.abc import Callable
from dataclasses import dataclass

from openpyxl.formula import Tokenizer
from openpyxl.formula.tokenizer import Token, TokenizerError
from openpyxl.utils.cell import column_index_from_string

from excel_mcp.calc.values import ERRORS, REF, UncalculableError
from excel_mcp.structured import StructuredRef, parse_structured

MAX_NESTING = 64  # Excel's limit on nested function calls
MAX_FORMULA_LENGTH = 8192


@dataclass(frozen=True)
class Literal:
    value: float | str | bool | None


@dataclass(frozen=True)
class ErrorLiteral:
    code: str


@dataclass(frozen=True)
class Ref:
    """A rectangle; bounds are None for whole-row or whole-column references."""

    sheets: tuple[str, ...]
    top: int | None
    left: int | None
    bottom: int | None
    right: int | None
    relative: bool


@dataclass(frozen=True)
class TableRef:
    """A part of a table; the engine replaces it by the cells it covers."""

    ref: StructuredRef


@dataclass(frozen=True)
class Name:
    sheet: str | None
    name: str


@dataclass(frozen=True)
class Unary:
    op: str
    operand: "Node"


@dataclass(frozen=True)
class Percent:
    operand: "Node"


@dataclass(frozen=True)
class Binary:
    op: str
    left: "Node"
    right: "Node"


@dataclass(frozen=True)
class Call:
    name: str
    args: tuple["Node", ...]
    prefixed: bool


@dataclass(frozen=True)
class ArrayLiteral:
    rows: tuple[tuple["Node", ...], ...]


Node = (
    Literal | ErrorLiteral | Ref | TableRef | Name | Unary | Percent | Binary | Call | ArrayLiteral
)

_PRECEDENCE = {
    "=": 1, "<>": 1, "<": 1, ">": 1, "<=": 1, ">=": 1,
    "&": 2,
    "+": 3, "-": 3,
    "*": 4, "/": 4,
    "^": 5,
}  # fmt: skip
_CELL = r"(\$?)([A-Za-z]{1,3})(\$?)(\d+)"
_RANGE = re.compile(rf"{_CELL}(:{_CELL})?")
_COLUMNS = re.compile(r"(\$?)([A-Za-z]{1,3}):(\$?)([A-Za-z]{1,3})")
_ROWS = re.compile(r"(\$?)(\d+):(\$?)(\d+)")
_FUNCTION_PREFIXES = ("_xlfn.", "_xlws.", "_xlpm.", "_xludf.", "_xleta.")


def parse(formula: str) -> Node:
    if len(formula) > MAX_FORMULA_LENGTH:
        raise UncalculableError("formula too long")
    try:
        tokens = [t for t in Tokenizer(formula).items if t.type != Token.WSPACE]
    except TokenizerError:
        raise UncalculableError("unparseable formula") from None
    parser = _Parser(tokens)
    node = parser.expression(0)
    if parser.position != len(tokens):
        raise UncalculableError("unparseable formula")
    return node


class _Parser:
    def __init__(self, tokens: list[Token]) -> None:
        self.tokens = tokens
        self.position = 0
        self.depth = 0

    def nested(self, build: Callable[[], Node]) -> Node:
        self.depth += 1
        if self.depth > MAX_NESTING:
            raise UncalculableError(f"formula nested more than {MAX_NESTING} levels")
        try:
            return build()
        finally:
            self.depth -= 1

    def peek(self) -> Token | None:
        return self.tokens[self.position] if self.position < len(self.tokens) else None

    def take(self) -> Token:
        token = self.peek()
        if token is None:
            raise UncalculableError("unparseable formula")
        self.position += 1
        return token

    def expression(self, minimum: int) -> Node:
        left = self.unary()
        while (token := self.peek()) and token.type == Token.OP_IN:
            precedence = _PRECEDENCE.get(token.value)
            if precedence is None:
                raise UncalculableError(f"operator {token.value}")
            if precedence < minimum:
                break
            self.take()
            left = Binary(token.value, left, self.expression(precedence + 1))
        return left

    def unary(self) -> Node:
        token = self.peek()
        if token and token.type == Token.OP_PRE:
            self.take()
            return self.nested(lambda: Unary(token.value, self.unary()))
        node = self.primary()
        while (token := self.peek()) and token.type == Token.OP_POST:
            self.take()
            node = Percent(node)
        return node

    def primary(self) -> Node:
        token = self.take()
        if token.type == Token.OPERAND:
            return _operand(token)
        if token.type == Token.FUNC and token.subtype == Token.OPEN:
            return self.nested(lambda: self.call(token))
        if token.type == Token.PAREN and token.subtype == Token.OPEN:
            return self.nested(self.parenthesized)
        if token.type == Token.ARRAY and token.subtype == Token.OPEN:
            return self.nested(self.array)
        raise UncalculableError("unparseable formula")

    def parenthesized(self) -> Node:
        node = self.expression(0)
        self.expect(Token.PAREN)
        return node

    def expect(self, kind: str) -> None:
        token = self.take()
        if token.type != kind or token.subtype != Token.CLOSE:
            raise UncalculableError("unparseable formula")

    def call(self, token: Token) -> Call:
        raw = token.value.removesuffix("(")
        args: list[Node] = []
        if (nxt := self.peek()) and nxt.type == Token.FUNC and nxt.subtype == Token.CLOSE:
            self.take()
        else:
            while True:
                args.append(self.argument())
                separator = self.take()
                if separator.type == Token.FUNC and separator.subtype == Token.CLOSE:
                    break
                if separator.type != Token.SEP or separator.subtype != Token.ARG:
                    raise UncalculableError("unparseable formula")
        return Call(normalize_name(raw), tuple(args), raw.lower().startswith(_FUNCTION_PREFIXES))

    def argument(self) -> Node:
        token = self.peek()
        if token and (
            token.type == Token.SEP or (token.type == Token.FUNC and token.subtype == Token.CLOSE)
        ):
            return Literal(None)
        return self.expression(0)

    def array(self) -> ArrayLiteral:
        rows: list[list[Node]] = [[]]
        while True:
            rows[-1].append(self.expression(0))
            token = self.take()
            if token.type == Token.ARRAY and token.subtype == Token.CLOSE:
                break
            if token.type == Token.SEP and token.subtype == Token.ROW:
                rows.append([])
            elif token.type != Token.SEP:
                raise UncalculableError("unparseable formula")
        return ArrayLiteral(tuple(tuple(row) for row in rows))


def normalize_name(raw: str) -> str:
    name = raw.upper()
    while name.lower().startswith(_FUNCTION_PREFIXES):
        name = name.split(".", 1)[1]
    return name


def _operand(token: Token) -> Node:
    text = token.value
    if token.subtype == Token.NUMBER:
        return Literal(float(text))
    if token.subtype == Token.TEXT:
        return Literal(text[1:-1].replace('""', '"'))
    if token.subtype == Token.LOGICAL:
        return Literal(text.upper() == "TRUE")
    if token.subtype == Token.ERROR:
        return ErrorLiteral(text.upper())
    return parse_reference(text)


def parse_reference(text: str) -> Node:
    if "[" in text:
        found = parse_structured(text)
        if found is None:
            raise UncalculableError("structured reference")
        return TableRef(found)
    qualifier, body = _split_sheet(text)
    if "#" in body.replace("#REF!", ""):
        raise UncalculableError("structured reference")
    if body.upper() == "#REF!":
        return ErrorLiteral(REF.code)
    sheets = _sheets(qualifier) if qualifier is not None else ()
    if match := _RANGE.fullmatch(body):
        first = match.groups()[:4]
        last = match.groups()[5:9] if match[5] else first
        return _ref(sheets, first, last)
    if match := _COLUMNS.fullmatch(body):
        left, right = match[2], match[4]
        relative = not (match[1] and match[3])
        return Ref(
            sheets,
            None,
            column_index_from_string(left.upper()),
            None,
            column_index_from_string(right.upper()),
            relative,
        )
    if match := _ROWS.fullmatch(body):
        relative = not (match[1] and match[3])
        return Ref(sheets, int(match[2]), None, int(match[4]), None, relative)
    if text.upper() in ERRORS:
        return ErrorLiteral(text.upper())
    return Name(sheets[0] if sheets else None, body.removeprefix("_xlpm."))


def _ref(sheets: tuple[str, ...], first: tuple[str, ...], last: tuple[str, ...]) -> Ref:
    relative = not all((first[0], first[2], last[0], last[2]))
    rows = sorted((int(first[3]), int(last[3])))
    cols = sorted(
        (
            column_index_from_string(first[1].upper()),
            column_index_from_string(last[1].upper()),
        )
    )
    return Ref(sheets, rows[0], cols[0], rows[1], cols[1], relative)


def _split_sheet(text: str) -> tuple[str | None, str]:
    quoted = False
    cut = -1
    for index, character in enumerate(text):
        if character == "'":
            quoted = not quoted
        elif character == "!" and not quoted:
            cut = index
    if cut < 0:
        return None, text
    return text[:cut], text[cut + 1 :]


def _sheets(qualifier: str) -> tuple[str, ...]:
    name = qualifier
    if len(name) > 1 and name.startswith("'") and name.endswith("'"):
        name = name[1:-1].replace("''", "'")
    if name.startswith("["):
        raise UncalculableError("reference to another workbook")
    return (name,)
