"""Parsing and formatting of A1-style cell references."""

from dataclasses import dataclass

from openpyxl.utils.cell import get_column_letter, range_boundaries

from excel_mcp.errors import InvalidArgumentError, LimitExceededError

MAX_ROW = 1_048_576
MAX_COLUMN = 16_384


@dataclass(frozen=True)
class CellRange:
    """A rectangular block of cells, 1-based and inclusive."""

    min_row: int
    min_col: int
    max_row: int
    max_col: int

    @property
    def rows(self) -> int:
        return self.max_row - self.min_row + 1

    @property
    def cols(self) -> int:
        return self.max_col - self.min_col + 1

    @property
    def size(self) -> int:
        return self.rows * self.cols

    @property
    def top_left(self) -> str:
        return cell_name(self.min_row, self.min_col)

    def overlaps(self, other: "CellRange") -> bool:
        return not (
            self.max_row < other.min_row
            or other.max_row < self.min_row
            or self.max_col < other.min_col
            or other.max_col < self.min_col
        )

    def within(self, max_cells: int) -> "CellRange":
        """Return the range, or raise if it has more than ``max_cells`` cells."""
        if self.size > max_cells:
            raise LimitExceededError(
                f"Range {self} has {self.size:,} cells; at most {max_cells:,} can be "
                "processed per call. Use a smaller range."
            )
        return self

    def __str__(self) -> str:
        start = self.top_left
        end = cell_name(self.max_row, self.max_col)
        return start if start == end else f"{start}:{end}"


def cell_name(row: int, col: int) -> str:
    return f"{get_column_letter(col)}{row}"


def parse_range(ref: str) -> CellRange:
    """Parse a bounded range such as ``B2`` or ``A1:C10`` (``$`` markers allowed)."""
    try:
        min_col, min_row, max_col, max_row = range_boundaries(ref.strip())
    except (ValueError, TypeError):
        raise InvalidArgumentError(
            f"Invalid range {ref!r}. Use A1 notation such as 'B2' or 'A1:C10'."
        ) from None
    if min_col is None or min_row is None or max_col is None or max_row is None:
        raise InvalidArgumentError(
            f"Range {ref!r} must have explicit rows and columns, such as 'A1:C10'."
        )
    cell_range = CellRange(
        min_row=min(min_row, max_row),
        min_col=min(min_col, max_col),
        max_row=max(min_row, max_row),
        max_col=max(min_col, max_col),
    )
    if cell_range.min_row < 1 or cell_range.max_row > MAX_ROW or cell_range.max_col > MAX_COLUMN:
        raise InvalidArgumentError(f"Range {ref!r} is outside the worksheet limits.")
    return cell_range


def parse_cell(ref: str) -> tuple[int, int]:
    """Parse a single cell such as ``B2`` into ``(row, column)``."""
    cell_range = parse_range(ref)
    if cell_range.size != 1:
        raise InvalidArgumentError(f"Expected a single cell such as 'B2', got {ref!r}.")
    return cell_range.min_row, cell_range.min_col
