"""Range transforms from Excel's Data and Fill commands."""

from typing import Literal

from openpyxl.worksheet.worksheet import Worksheet
from pydantic import Field

from excel_mcp.errors import InvalidArgumentError
from excel_mcp.inputs import InputModel
from excel_mcp.operations.comparison import FormulaResults
from excel_mcp.operations.duplicates import remove_duplicates
from excel_mcp.operations.fill import DateUnit, Direction, Series, fill_range
from excel_mcp.operations.text_split import text_to_columns
from excel_mcp.refs import CellRange

Operation = Literal["remove_duplicates", "text_to_columns", "fill"]

_FIELDS = {
    "remove_duplicates": {"columns", "has_header"},
    "text_to_columns": {"delimiters", "fixed_widths_chars", "text_qualifier", "merge_delimiters"},
    "fill": {"direction", "series", "step", "stop", "unit"},
}


class Transform(InputModel):
    """Only the fields named for the ``operation`` apply."""

    operation: Operation = Field(
        description="remove_duplicates keeps the first of equal rows and moves the rest of the "
        "range up. text_to_columns splits one column into the columns to its right. fill fills "
        "the range from its first row or column."
    )
    columns: list[str] | None = Field(
        default=None,
        description="remove_duplicates: columns that must match (header text or letter). "
        "Default: all.",
    )
    has_header: bool = Field(default=True, description="remove_duplicates: first row is a header.")
    delimiters: list[str] | None = Field(
        default=None,
        description="text_to_columns: 'tab', 'semicolon', 'comma', 'space' or a character.",
    )
    fixed_widths_chars: list[int] | None = Field(
        default=None,
        description="text_to_columns instead of delimiters: widths of all fields but the last.",
    )
    text_qualifier: Literal['"', "'", ""] = Field(
        default='"', description="text_to_columns: quote that protects delimiters; '' for none."
    )
    merge_delimiters: bool = Field(
        default=False, description="text_to_columns: consecutive delimiters count as one."
    )
    direction: Direction | None = Field(
        default=None, description="fill: down from the first row, or right from the first column."
    )
    series: Series = Field(
        default="copy",
        description="fill: copy repeats the first line; the others continue each seed cell.",
    )
    step: float = Field(
        default=1, description="fill series: amount added (linear, date) or multiplied by (growth)."
    )
    stop: str | float | None = Field(
        default=None, description="fill series: last value; dates as '2026-12-31'."
    )
    unit: DateUnit = Field(default="day", description="fill date series: step unit.")


def apply_transform(
    sheet: Worksheet,
    area: CellRange,
    spec: Transform,
    results: FormulaResults,
    max_cells: int,
) -> tuple[CellRange, str | None]:
    """Run the transform; return the cells it changed and, if it counted something, a note."""
    spec.reject_unused(_FIELDS[spec.operation] | {"operation"}, spec.operation)
    match spec.operation:
        case "remove_duplicates":
            removed = remove_duplicates(sheet, area, spec.columns, spec.has_header, results)
            return area, f"Removed {removed} duplicate rows."
        case "text_to_columns":
            split = text_to_columns(
                sheet,
                area,
                spec.delimiters,
                spec.fixed_widths_chars,
                spec.text_qualifier,
                spec.merge_delimiters,
                max_cells,
            )
            return area, f"Split {split} cells."
        case "fill":
            if spec.direction is None:
                raise InvalidArgumentError("fill needs a direction, 'down' or 'right'.")
            filled = fill_range(
                sheet, area, spec.direction, spec.series, spec.step, spec.stop, spec.unit, max_cells
            )
            return filled, None
