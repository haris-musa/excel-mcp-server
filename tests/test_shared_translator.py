"""A shared formula is translated as openpyxl's Translator does, for every formula."""

import random

import pytest
from openpyxl.formula.tokenizer import TokenizerError
from openpyxl.formula.translate import Translator, TranslatorError
from openpyxl.utils import get_column_letter

from excel_mcp.shared_translator import SharedTranslator

REFERENCES = [
    "A1", "$A$1", "$A1", "A$1", "b7", "$c$9", "XFD1048576", "AA10", "ZZZ1", "Z9",
    "A1:B5", "$A$1:B$2", "A1:$B$2", "a2:C3", "3:5", "$3:$5", "3:$5", "A:C", "$A:B", "A:$C",
    "Sheet1!A1", "'My Sheet'!$B$2:C3", "Sheet1!A:A", "Sheet1!2:3", "'It''s'!A1",
    "Sheet1:Sheet3!A1", "Total", "tax_rate", "Table1[Col]", "Table1[[#This Row],[Col]]",
    "Name:C2", "C2:Name", "A1:Name:B2", "R1C1", "A1B", "Sheet1!Name", "[1]Sheet1!A1",
    "A1#", "INDIRECT", "1048576:1048577",
]  # fmt: skip
OPERANDS = ['"A1 and B2"', '"say ""B2"""', "1", "2.5E3", "TRUE", "#REF!", "{1,2;3,4}", "-3", "50%"]
OPERATORS = ["+", "-", "*", "/", "^", "&", "=", "<>", ">=", "<", " + ", "  *  ", ",", ";"]
FUNCTIONS = ["SUM", "IF", "VLOOKUP", "INDEX", "AND", "TEXT", "_xlfn.XLOOKUP", "MAX"]


def formula(random_: random.Random, depth: int = 0) -> str:
    parts = []
    for _ in range(random_.randint(1, 6)):
        kind = random_.random()
        if kind < 0.5:
            parts.append(random_.choice(REFERENCES))
        elif kind < 0.65:
            parts.append(random_.choice(OPERANDS))
        elif kind < 0.9 and depth < 2:
            arguments = ",".join(formula(random_, depth + 1) for _ in range(random_.randint(1, 3)))
            parts.append(f"{random_.choice(FUNCTIONS)}({arguments})")
        else:
            parts.append(f"({formula(random_, depth + 1)})" if depth < 2 else "1")
        parts.append(random_.choice(OPERATORS))
    return parts[0] if len(parts) == 1 else "".join(parts[:-1])


def outcome(make: type[Translator], text: str, origin: str, dest: str) -> str:
    try:
        return make(text, origin).translate_formula(dest)
    except (TranslatorError, TokenizerError) as error:
        return type(error).__name__


@pytest.mark.parametrize("seed", range(8))
def test_every_formula_translates_as_openpyxl_translates_it(seed: int) -> None:
    random_ = random.Random(seed)
    for _ in range(250):
        text = "=" + formula(random_)
        origin = f"{get_column_letter(random_.randint(1, 40))}{random_.randint(1, 60)}"
        for _ in range(6):
            dest = f"{get_column_letter(random_.randint(1, 60))}{random_.randint(1, 90)}"
            assert outcome(SharedTranslator, text, origin, dest) == outcome(
                Translator, text, origin, dest
            ), (text, origin, dest)


@pytest.mark.parametrize(
    "text", ["", "=", "text", "=  ", "= A1", "=A1 ", "\n=A1", "=A1+\nB1", "=SUM(\n A1,\n B1 )"]
)
def test_odd_formulas_translate_as_openpyxl_translates_them(text: str) -> None:
    for dest in ("A1", "B2", "C9", "AB40"):
        assert outcome(SharedTranslator, text, "B2", dest) == outcome(Translator, text, "B2", dest)


def test_offsets_that_leave_the_sheet_are_refused_as_they_are() -> None:
    for text, dest in [
        ("=A1", "A1"),
        ("=B1", "A2"),
        ("=ZZZ1", "C2"),
        ("=1:2", "B1"),
        ("=$A1+A$1", "A1"),
    ]:
        assert outcome(SharedTranslator, text, "B2", dest) == outcome(Translator, text, "B2", dest)
    with pytest.raises(TranslatorError):
        SharedTranslator("=A1", "B2").translate_formula("B1")


def test_row_and_column_offsets_can_be_given_instead_of_a_cell() -> None:
    translator = SharedTranslator("=A1+$B$2+C$3:$D4", "A1")
    assert translator.translate_formula(row_delta=2, col_delta=3) == Translator(
        "=A1+$B$2+C$3:$D4", "A1"
    ).translate_formula(row_delta=2, col_delta=3)
