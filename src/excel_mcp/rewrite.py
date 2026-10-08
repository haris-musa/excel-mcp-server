"""Rewriting the references inside formulas, shared by sheet copies and row or column edits.

Formulas are tokenized with openpyxl's Excel tokenizer and only the reference
tokens are rewritten, so text, names and functions are never touched.
"""

from abc import ABC, abstractmethod
from collections.abc import Iterator
from typing import Any

from openpyxl.chart.data_source import MultiLevelStrRef, NumRef, StrRef
from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.formula import Tokenizer
from openpyxl.formula.tokenizer import Token


class ReferenceRewriter(ABC):
    """Rewrites each reference of a formula; ``host`` is the sheet the formula is written on."""

    @abstractmethod
    def reference(self, text: str, host: str) -> str:
        """Rewrite one reference such as ``Sheet1!$A$1:B5``; names stay as they are."""

    def formula(self, formula: str, host: str) -> str:
        """Rewrite a formula that starts with "=" ."""
        tokenizer = Tokenizer(formula)
        for token in tokenizer.items:
            if token.subtype == Token.RANGE:
                token.value = self.reference(token.value, host)
            elif token.type == Token.FUNC and token.subtype == Token.OPEN and ":" in token.value:
                # A reference can start a range that ends in a function, as in "A1:INDEX(".
                head, colon, function = token.value.rpartition(":")
                token.value = self.reference(head, host) + colon + function
        return tokenizer.render()

    def operand(self, operand: str, host: str) -> str:
        """`formula` for text stored without the "=", as in names, rules and chart references."""
        return self.formula(f"={operand}", host).removeprefix("=")


def chart_references(node: Any) -> Iterator[NumRef | StrRef | MultiLevelStrRef]:
    """Every cell reference inside a chart: series values, categories, titles."""
    if isinstance(node, NumRef | StrRef | MultiLevelStrRef):
        yield node
    elif isinstance(node, Serialisable):
        for field, value in vars(node).items():
            if field != "_charts":
                yield from chart_references(value)
    elif isinstance(node, list | tuple):
        for item in node:
            yield from chart_references(item)
