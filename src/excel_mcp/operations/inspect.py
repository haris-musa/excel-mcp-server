"""Read-only descriptions of workbooks, sheets and folders."""

from pathlib import Path

from openpyxl.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet
from pydantic import BaseModel

from excel_mcp.operations.cells import streamed_used_range, used_range
from excel_mcp.operations.chart_index import list_charts
from excel_mcp.operations.chart_info import ChartInfo
from excel_mcp.operations.conditional_x14 import extended_rules
from excel_mcp.operations.hyperlinks import LinkInfo, list_links
from excel_mcp.operations.images import ImageInfo, list_images
from excel_mcp.operations.layout import hidden_lines
from excel_mcp.operations.names import DefinedNameInfo, list_defined_names
from excel_mcp.operations.notes import NoteInfo, list_notes
from excel_mcp.operations.pivot_index import PivotInfo, list_pivots
from excel_mcp.operations.sheet_view import ViewInfo, read_view
from excel_mcp.operations.sparkline_style import SparklineInfo
from excel_mcp.operations.sparklines import list_sparklines
from excel_mcp.operations.workbook_settings import (
    CalculationInfo,
    PropertiesInfo,
    read_calculation,
    read_properties,
    structure_protected,
)
from excel_mcp.paths import EXCEL_SUFFIXES
from excel_mcp.workspace import streamed_worksheets


class SheetSummary(BaseModel):
    name: str
    used_range: str
    hidden: bool = False
    active: bool = False


class WorkbookInfo(BaseModel):
    sheets: list[SheetSummary]
    chart_sheets: list[str] = []
    defined_names: list[DefinedNameInfo] = []
    has_vba: bool = False
    doc_properties: PropertiesInfo = PropertiesInfo()
    calculation: CalculationInfo = CalculationInfo()
    structure_protected: bool = False


class DataValidationInfo(BaseModel):
    range: str
    type: str | None = None
    operator: str | None = None
    formula1: str | None = None
    formula2: str | None = None


class ConditionalFormatInfo(BaseModel):
    range: str
    type: str


def list_conditional_formats(sheet: Worksheet) -> list[ConditionalFormatInfo]:
    rules = [
        ConditionalFormatInfo(range=str(entry.sqref), type=rule.type)
        for entry in sheet.conditional_formatting
        for rule in entry.rules
    ]
    return rules + [ConditionalFormatInfo(range=r, type=t) for r, t in extended_rules(sheet)]


class SheetDetails(BaseModel):
    used_range: str
    freeze_panes: str | None = None
    auto_filter: str | None = None
    merged_ranges: list[str] = []
    notes: list[NoteInfo] = []
    tables: dict[str, str] = {}
    charts: list[ChartInfo] = []
    pivot_tables: list[PivotInfo] = []
    data_validations: list[DataValidationInfo] = []
    conditional_formats: list[ConditionalFormatInfo] = []
    sparklines: list[SparklineInfo] = []
    column_widths: dict[str, float] = {}
    hidden_rows: list[str] = []
    hidden_columns: list[str] = []
    images: list[ImageInfo] = []
    hyperlinks: dict[str, LinkInfo] = {}
    print_area: str | None = None
    protected: bool = False
    view: ViewInfo = ViewInfo()


def describe_workbook(workbook: Workbook, has_vba: bool, company: str) -> WorkbookInfo:
    """Summarize a streamed workbook: each sheet costs one pass over its cells."""
    return WorkbookInfo(
        sheets=[
            SheetSummary(
                name=sheet.title,
                used_range=str(streamed_used_range(sheet)),
                hidden=sheet.sheet_state != "visible",
                active=workbook.active is sheet,
            )
            for sheet in streamed_worksheets(workbook)
        ],
        chart_sheets=[sheet.title for sheet in workbook.chartsheets],
        defined_names=list_defined_names(workbook),
        has_vba=has_vba,
        doc_properties=read_properties(workbook, company),
        calculation=read_calculation(workbook),
        structure_protected=structure_protected(workbook),
    )


def describe_sheet(sheet: Worksheet) -> SheetDetails:
    return SheetDetails(
        used_range=str(used_range(sheet)),
        freeze_panes=sheet.freeze_panes,
        auto_filter=sheet.auto_filter.ref,
        merged_ranges=sorted(str(merged) for merged in sheet.merged_cells.ranges),
        notes=list_notes(sheet),
        tables=dict(sheet.tables.items()),
        charts=list_charts(sheet),
        pivot_tables=list_pivots(sheet),
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
        conditional_formats=list_conditional_formats(sheet),
        sparklines=list_sparklines(sheet),
        column_widths={
            letter: dimension.width
            for letter, dimension in sorted(sheet.column_dimensions.items())
            if dimension.customWidth
        },
        hidden_rows=hidden_lines(sheet, "rows"),
        hidden_columns=hidden_lines(sheet, "columns"),
        images=list_images(sheet),
        hyperlinks=list_links(sheet),
        print_area=sheet.print_area or None,
        protected=bool(sheet.protection.sheet),
        view=read_view(sheet),
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
