"""Clearing conditional formats and data validation from cells, as Excel's Clear Rules and
Data Validation > Clear All do: a rule loses the cells and keeps the rest of its range."""

from copy import copy
from typing import cast

from openpyxl.formatting.formatting import ConditionalFormattingList
from openpyxl.workbook import Workbook
from openpyxl.worksheet.cell_range import MultiCellRange
from openpyxl.worksheet.worksheet import Worksheet

from excel_mcp.operations.shifting import Mover
from excel_mcp.package import extensions, state_of
from excel_mcp.refs import CellRange, clip_areas


def clear_conditional_formats(sheet: Worksheet, hole: CellRange) -> None:
    package = state_of(cast(Workbook, sheet.parent)).sheet(sheet)
    host = sheet.title
    rules = ConditionalFormattingList()
    rule_extensions = {}
    for entry in sheet.conditional_formatting:
        old = str(entry.sqref)
        found = _clip(entry.sqref, hole)
        for rule in entry.rules:
            kept = package.rule_extensions.pop((old, str(rule.priority)), None)
            if found is None:
                continue
            sqref, mover = found
            clone = copy(rule)
            clone.formula = [mover.operand(f, host) for f in rule.formula or []]
            rules.add(sqref, clone)
            if kept is not None:
                rule_extensions[(sqref, str(rule.priority))] = kept
    package.rule_extensions.update(rule_extensions)
    sheet.conditional_formatting = rules
    extensions.clip_rules(package.extensions, extensions.CONDITIONAL_FORMATS, hole)


def clear_validation(sheet: Worksheet, hole: CellRange) -> None:
    package = state_of(cast(Workbook, sheet.parent)).sheet(sheet)
    host = sheet.title
    validations = []
    for validation in sheet.data_validations.dataValidation:
        found = _clip(validation.sqref, hole)
        if found is None:
            continue
        validation = copy(validation)
        validation.sqref = MultiCellRange(found[0])
        for field in ("formula1", "formula2"):
            if formula := getattr(validation, field):
                setattr(validation, field, found[1].operand(formula, host))
        validations.append(validation)
    sheet.data_validations.dataValidation = validations
    extensions.clip_rules(package.extensions, extensions.DATA_VALIDATIONS, hole)


def _clip(sqref: MultiCellRange, hole: CellRange) -> tuple[str, Mover] | None:
    """The range text without ``hole`` and the mover for the formulas written for its corner."""
    found = clip_areas([CellRange.of(r) for r in sqref.ranges], hole)
    return None if found is None else (" ".join(map(str, found[0])), Mover(found[1], found[2]))
