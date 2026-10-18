"""Comparisons shared by the audit recorder and its test."""

from openpyxl.utils.cell import range_boundaries

CLIP_ROWS = 60  # whole columns are compared over their first rows
CLIP_COLUMNS = 40


def cell_set(ranges: list[str]) -> set[tuple[int, int]]:
    """The cells of range texts such as 'D2:D9', '$F$1', 'B:B' or '2:3', within the clip."""
    cells: set[tuple[int, int]] = set()
    for text in ranges:
        if not text:
            continue
        left, top, right, bottom = range_boundaries(text.replace("$", ""))
        for row in range(top or 1, min(bottom or CLIP_ROWS, CLIP_ROWS) + 1):
            for col in range(left or 1, min(right or CLIP_COLUMNS, CLIP_COLUMNS) + 1):
                cells.add((row, col))
    return cells
