"""Page setup for printing: orientation, paper, scaling, margins, titles, headers."""

from typing import Literal

from openpyxl.worksheet.properties import PageSetupProperties
from openpyxl.worksheet.worksheet import Worksheet
from pydantic import BaseModel, Field, model_validator

from excel_mcp.operations.spans import Axis, format_span, parse_span
from excel_mcp.refs import parse_range

CM_PER_INCH = 2.54
PAPER_SIZES = {"letter": 1, "legal": 5, "tabloid": 3, "a3": 8, "a4": 9, "a5": 11}


class FitToPages(BaseModel):
    """Shrink the printout to this many pages."""

    wide: int = Field(default=1, ge=0, le=100, description="Pages across; 0 = as many as needed.")
    tall: int = Field(default=1, ge=0, le=100, description="Pages down; 0 = as many as needed.")


class Margins(BaseModel):
    """Page margins in centimetres. Fields left as null are not changed."""

    left: float | None = Field(default=None, ge=0, le=50, description="Left margin in cm.")
    right: float | None = Field(default=None, ge=0, le=50, description="Right margin in cm.")
    top: float | None = Field(default=None, ge=0, le=50, description="Top margin in cm.")
    bottom: float | None = Field(default=None, ge=0, le=50, description="Bottom margin in cm.")
    header: float | None = Field(default=None, ge=0, le=50, description="Header distance in cm.")
    footer: float | None = Field(default=None, ge=0, le=50, description="Footer distance in cm.")


class HeaderFooterText(BaseModel):
    """Text in the left, center and right of a header or footer. Fields left as null are kept."""

    left: str | None = Field(default=None, description="Left text; '' clears it.")
    center: str | None = Field(default=None, description="Center text; '' clears it.")
    right: str | None = Field(default=None, description="Right text; '' clears it.")


class PrintSetup(BaseModel):
    """Print settings. Fields left as null are not changed."""

    orientation: Literal["portrait", "landscape"] | None = Field(
        default=None, description="Page orientation."
    )
    paper_size: Literal["letter", "legal", "tabloid", "a3", "a4", "a5"] | None = Field(
        default=None, description="Paper size."
    )
    scale: int | None = Field(
        default=None, ge=10, le=400, description="Print at this percent; not with fit_to_pages."
    )
    fit_to_pages: FitToPages | None = Field(
        default=None, description="Fit the printout to a number of pages."
    )
    margins_cm: Margins | None = Field(default=None, description="Page margins.")
    print_area: str | None = Field(
        default=None, description="Range to print, e.g. 'A1:H40'; '' prints the whole sheet."
    )
    title_rows: str | None = Field(
        default=None, description="Rows repeated on every page, e.g. '1:2'; '' clears."
    )
    title_columns: str | None = Field(
        default=None, description="Columns repeated on every page, e.g. 'A:B'; '' clears."
    )
    center_horizontally: bool | None = Field(
        default=None, description="Center the data between the left and right margins."
    )
    center_vertically: bool | None = Field(
        default=None, description="Center the data between the top and bottom margins."
    )
    gridlines: bool | None = Field(default=None, description="Print cell gridlines.")
    header: HeaderFooterText | None = Field(
        default=None,
        description="Page header. Codes: &P page, &N page count, &D date, &A sheet, &F file.",
    )
    footer: HeaderFooterText | None = Field(default=None, description="Page footer, same codes.")

    @model_validator(mode="after")
    def _scale_or_fit(self) -> "PrintSetup":
        if self.scale is not None and self.fit_to_pages is not None:
            raise ValueError("Set either scale or fit_to_pages, not both.")
        return self


def apply_print_setup(sheet: Worksheet, setup: PrintSetup) -> None:
    page = sheet.page_setup
    if setup.orientation is not None:
        page.orientation = setup.orientation
    if setup.paper_size is not None:
        page.paperSize = PAPER_SIZES[setup.paper_size]
    if setup.scale is not None:
        _fit_to_page(sheet, False)
        page.scale = setup.scale
    if setup.fit_to_pages is not None:
        _fit_to_page(sheet, True)
        page.fitToWidth = setup.fit_to_pages.wide
        page.fitToHeight = setup.fit_to_pages.tall
    if setup.margins_cm is not None:
        _set_margins(sheet, setup.margins_cm)
    if setup.print_area is not None:
        sheet.print_area = str(parse_range(setup.print_area)) if setup.print_area else None
    # openpyxl's setters ignore None, so clearing titles needs the attributes themselves.
    if setup.title_rows is not None:
        sheet._print_rows = None  # pyright: ignore[reportAttributeAccessIssue]
        sheet.print_title_rows = _title(setup.title_rows, "rows")
    if setup.title_columns is not None:
        sheet._print_cols = None  # pyright: ignore[reportAttributeAccessIssue]
        sheet.print_title_cols = _title(setup.title_columns, "columns")
    options = sheet.print_options
    if setup.center_horizontally is not None:
        options.horizontalCentered = setup.center_horizontally
    if setup.center_vertically is not None:
        options.verticalCentered = setup.center_vertically
    if setup.gridlines is not None:
        options.gridLines = setup.gridlines
    _set_text(sheet.oddHeader, setup.header)
    _set_text(sheet.oddFooter, setup.footer)


def _fit_to_page(sheet: Worksheet, enabled: bool) -> None:
    if sheet.sheet_properties.pageSetUpPr is None:
        sheet.sheet_properties.pageSetUpPr = PageSetupProperties()
    sheet.sheet_properties.pageSetUpPr.fitToPage = enabled


def _set_margins(sheet: Worksheet, margins: Margins) -> None:
    for name, value in margins.model_dump(exclude_none=True).items():
        setattr(sheet.page_margins, name, round(value / CM_PER_INCH, 4))


def _set_text(part, text: HeaderFooterText | None) -> None:
    if text is None:
        return
    for position, value in text.model_dump(exclude_none=True).items():
        getattr(part, position).text = value or None


def _title(text: str, axis: Axis) -> str | None:
    if not text.strip():
        return None
    first, last = parse_span(text, axis)
    return format_span(first, last, axis)
