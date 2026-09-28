"""Read-only descriptions of workbooks, sheets and folders."""

from pathlib import Path

from openpyxl.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet
from pydantic import BaseModel

from excel_mcp.operations.cells import used_range
from excel_mcp.paths import EXCEL_SUFFIXES
from excel_mcp.workspace import worksheets


class SheetSummary(BaseModel):
    name: str
    used_range: str
    rows: int
    columns: int
    visible: bool


class WorkbookInfo(BaseModel):
    path: str
    size_bytes: int
    has_vba: bool
    sheets: list[SheetSummary]
    defined_names: list[str]


class NamedRange(BaseModel):
    name: str
    range: str


class DataValidationInfo(BaseModel):
    range: str
    type: str | None
    operator: str | None
    formula1: str | None
    formula2: str | None


class ConditionalFormatInfo(BaseModel):
    range: str
    type: str


class SheetDetails(BaseModel):
    name: str
    used_range: str
    freeze_panes: str | None
    auto_filter: str | None
    merged_ranges: list[str]
    tables: list[NamedRange]
    chart_count: int
    data_validations: list[DataValidationInfo]
    conditional_formats: list[ConditionalFormatInfo]
    column_widths: dict[str, float]


class WorkbookFile(BaseModel):
    path: str
    size_bytes: int


def summarize_sheet(sheet: Worksheet) -> SheetSummary:
    area = used_range(sheet)
    return SheetSummary(
        name=sheet.title,
        used_range=str(area),
        rows=area.rows,
        columns=area.cols,
        visible=sheet.sheet_state == "visible",
    )


def describe_workbook(
    workbook: Workbook, display_path: str, size_bytes: int, has_vba: bool
) -> WorkbookInfo:
    return WorkbookInfo(
        path=display_path,
        size_bytes=size_bytes,
        has_vba=has_vba,
        sheets=[summarize_sheet(sheet) for sheet in worksheets(workbook)],
        defined_names=sorted(workbook.defined_names),
    )


def describe_sheet(sheet: Worksheet) -> SheetDetails:
    return SheetDetails(
        name=sheet.title,
        used_range=str(used_range(sheet)),
        freeze_panes=sheet.freeze_panes,
        auto_filter=sheet.auto_filter.ref,
        merged_ranges=sorted(str(merged) for merged in sheet.merged_cells.ranges),
        tables=[NamedRange(name=name, range=ref) for name, ref in sheet.tables.items()],
        chart_count=len(sheet._charts),  # pyright: ignore[reportAttributeAccessIssue]
        data_validations=[
            DataValidationInfo(
                range=str(rule.sqref),
                type=rule.type,
                operator=rule.operator,
                formula1=rule.formula1,
                formula2=rule.formula2,
            )
            for rule in sheet.data_validations.dataValidation
        ],
        conditional_formats=[
            ConditionalFormatInfo(range=str(entry.sqref), type=rule.type)
            for entry in sheet.conditional_formatting
            for rule in entry.rules
        ],
        column_widths={
            letter: dimension.width
            for letter, dimension in sorted(sheet.column_dimensions.items())
            if dimension.customWidth
        },
    )


def list_workbooks(directory: Path, recursive: bool, limit: int) -> list[Path]:
    pattern = "**/*" if recursive else "*"
    found = (
        path
        for path in directory.glob(pattern)
        if path.suffix.lower() in EXCEL_SUFFIXES
        and path.is_file()
        and not path.name.startswith(("~$", ".~"))
    )
    return sorted(found)[:limit]
