"""Applying a `ReferenceRewriter` to every formula-bearing part of a workbook."""

from collections.abc import Iterator

from openpyxl.formatting.rule import FormatObject, Rule
from openpyxl.workbook import Workbook
from openpyxl.worksheet.formula import ArrayFormula

from excel_mcp.formulas import storable_formula, storable_operand
from excel_mcp.operations.cells import stored_cells
from excel_mcp.rewrite import ReferenceRewriter, chart_references
from excel_mcp.workspace import worksheets


def rewrite_formulas(workbook: Workbook, rewriter: ReferenceRewriter, names: list[str]) -> None:
    for sheet in worksheets(workbook):
        for cell in stored_cells(sheet):
            value = cell.value
            if isinstance(value, ArrayFormula):
                moved = rewriter.formula(str(value.text), sheet.title)
                if moved != value.text:
                    value.text = storable_formula(moved, names)
            elif cell.data_type == "f" and isinstance(value, str):
                moved = rewriter.formula(value, sheet.title)
                if moved != value:
                    cell.value = storable_formula(moved, names)
        for table in sheet.tables.values():
            for column in table.tableColumns:
                for formula in (column.calculatedColumnFormula, column.totalsRowFormula):
                    if formula is not None and formula.attr_text:
                        moved = rewriter.operand(formula.attr_text, sheet.title)
                        if moved != formula.attr_text:
                            formula.attr_text = storable_operand(moved, names)


def rewrite_names(workbook: Workbook, rewriter: ReferenceRewriter, names: list[str]) -> None:
    scopes = [("", workbook.defined_names)]
    scopes += [(sheet.title, sheet.defined_names) for sheet in worksheets(workbook)]
    for host, defined_names in scopes:
        for defined in defined_names.values():
            if defined.attr_text:
                moved = rewriter.operand(defined.attr_text, host)
                if moved != defined.attr_text:
                    defined.attr_text = storable_operand(moved, names)


def rewrite_charts(workbook: Workbook, rewriter: ReferenceRewriter, names: list[str]) -> None:
    """Series, category and title references of every chart in the workbook."""
    for sheet in (*workbook.worksheets, *workbook.chartsheets):
        for chart in sheet._charts:  # pyright: ignore[reportAttributeAccessIssue]
            references = {id(r): r for plot in chart._charts for r in chart_references(plot)}
            for reference in references.values():
                text = str(reference.f)
                moved = rewriter.operand(text, sheet.title)
                if moved != text:
                    reference.f = storable_operand(moved, names)


def rewrite_rules(workbook: Workbook, rewriter: ReferenceRewriter, names: list[str]) -> None:
    """Formulas of conditional formats and data validation, as written."""
    for sheet in worksheets(workbook):
        for entry in sheet.conditional_formatting:
            for rule in entry.rules:
                rule.formula = [
                    rewritten_operand(f, sheet.title, rewriter, names) for f in rule.formula or []
                ]
        for validation in sheet.data_validations.dataValidation:
            if validation.formula1:
                validation.formula1 = rewritten_operand(
                    validation.formula1, sheet.title, rewriter, names
                )
            if validation.formula2:
                validation.formula2 = rewritten_operand(
                    validation.formula2, sheet.title, rewriter, names
                )


def gated_operand(text: str, host: str, rewriter: ReferenceRewriter, names: list[str]) -> str:
    """``rewritten_operand`` under the argument order of the package layer's callback."""
    return rewritten_operand(text, host, rewriter, names)


def rewritten_operand(text: str, host: str, rewriter: ReferenceRewriter, names: list[str]) -> str:
    moved = rewriter.operand(text, host)
    return text if moved == text else storable_operand(moved, names)


def rule_points(rule: Rule) -> Iterator[FormatObject]:
    """The thresholds of a scale, data bar or icon set that are formulas."""
    for part in (rule.colorScale, rule.dataBar, rule.iconSet):
        for point in getattr(part, "cfvo", None) or []:
            if point.type == "formula" and point.val is not None:
                yield point
