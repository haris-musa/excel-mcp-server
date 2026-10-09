"""Workbook-wide settings: document properties, calculation options, structure protection."""

from typing import Literal

from openpyxl.workbook import Workbook
from openpyxl.workbook.properties import CalcProperties
from openpyxl.workbook.protection import WorkbookProtection
from pydantic import BaseModel, Field

from excel_mcp.errors import InvalidArgumentError
from excel_mcp.inputs import InputModel
from excel_mcp.operations.protection import password_matches
from excel_mcp.package import state_of

CalcMode = Literal["auto", "manual", "auto_except_tables"]
StoredMode = Literal["auto", "manual", "autoNoTable"]
_MODES: dict[CalcMode, StoredMode] = {
    "auto": "auto",
    "manual": "manual",
    "auto_except_tables": "autoNoTable",
}
_MODE_NAMES: dict[StoredMode, CalcMode] = {stored: mode for mode, stored in _MODES.items()}


class Properties(InputModel):
    """Fields left out are not changed; '' clears one."""

    title: str | None = None
    subject: str | None = None
    author: str | None = None
    keywords: str | None = None
    company: str | None = None


class Calculation(InputModel):
    mode: CalcMode | None = Field(
        default=None,
        description="'manual': only on request (F9); 'auto_except_tables': skips data tables.",
    )
    iterative: bool | None = Field(default=None, description="Allow circular references.")
    max_iterations: int | None = Field(default=None, ge=1, le=32_767)
    max_change: float | None = Field(
        default=None, gt=0, description="Stop when results change less than this."
    )
    full_calc_on_load: bool | None = Field(default=None, description="Recalculate on open.")


class StructureProtection(InputModel):
    enabled: bool = Field(description="False unprotects.")
    password: str | None = Field(
        default=None, max_length=255, description="To set; to unprotect, the current one."
    )


class WorkbookSettings(InputModel):
    doc_properties: Properties | None = None
    calculation: Calculation | None = None
    structure_protection: StructureProtection | None = Field(default=None)


class PropertiesInfo(BaseModel):
    title: str = ""
    subject: str = ""
    author: str = ""
    keywords: str = ""
    company: str = ""


class CalculationInfo(BaseModel):
    mode: CalcMode = "auto"
    iterative: bool = False
    max_iterations: int = 100
    max_change: float = 0.001
    full_calc_on_load: bool = False


def apply_settings(workbook: Workbook, settings: WorkbookSettings) -> None:
    if settings.doc_properties is not None:
        _set_properties(workbook, settings.doc_properties)
    if settings.calculation is not None:
        _set_calculation(workbook, settings.calculation)
    if settings.structure_protection is not None:
        _set_structure_protection(workbook, settings.structure_protection)


def read_properties(workbook: Workbook, company: str) -> PropertiesInfo:
    properties = workbook.properties
    return PropertiesInfo(
        title=properties.title or "",
        subject=properties.subject or "",
        author=properties.creator or "",
        keywords=properties.keywords or "",
        company=company,
    )


def read_calculation(workbook: Workbook) -> CalculationInfo:
    calc = workbook.calculation
    if calc is None:
        return CalculationInfo()
    return CalculationInfo(
        mode=_MODE_NAMES[calc.calcMode or "auto"],
        iterative=bool(calc.iterate),
        max_iterations=calc.iterateCount or 100,
        max_change=calc.iterateDelta or 0.001,
        full_calc_on_load=bool(calc.fullCalcOnLoad),
    )


def structure_protected(workbook: Workbook) -> bool:
    return bool(workbook.security and workbook.security.lockStructure)


def _set_properties(workbook: Workbook, change: Properties) -> None:
    properties = workbook.properties
    if change.title is not None:
        properties.title = change.title
    if change.subject is not None:
        properties.subject = change.subject
    if change.author is not None:
        properties.creator = change.author
    if change.keywords is not None:
        properties.keywords = change.keywords
    if change.company is not None:
        state_of(workbook).workbook.company = change.company


def _set_calculation(workbook: Workbook, change: Calculation) -> None:
    calc = workbook.calculation or CalcProperties(fullCalcOnLoad=None)
    if change.mode is not None:
        calc.calcMode = _MODES[change.mode]
    if change.iterative is not None:
        calc.iterate = change.iterative
    if change.max_iterations is not None:
        calc.iterateCount = change.max_iterations
    if change.max_change is not None:
        calc.iterateDelta = change.max_change
    if change.full_calc_on_load is not None:
        calc.fullCalcOnLoad = change.full_calc_on_load
    workbook.calculation = calc


def _set_structure_protection(workbook: Workbook, change: StructureProtection) -> None:
    current = workbook.security or WorkbookProtection()
    secured = bool(
        current.lockStructure and (current.workbookPassword or current.workbookHashValue)
    )
    if change.enabled:
        if secured:
            raise InvalidArgumentError(
                "The workbook structure is already protected with a password. Unprotect it first."
            )
        protection = WorkbookProtection(lockStructure=True)
        if change.password:
            protection.set_workbook_password(change.password)
        workbook.security = protection
        return
    if secured and not password_matches(
        change.password,
        legacy=current.workbookPassword,
        algorithm=current.workbookAlgorithmName,
        salt=current.workbookSaltValue,
        spin_count=current.workbookSpinCount,
        digest=current.workbookHashValue,
    ):
        raise InvalidArgumentError(
            "The workbook structure is password protected: give the right password."
        )
    workbook.security = WorkbookProtection()
