"""Parameter types shared by the tools, with the descriptions clients show the model."""

from typing import Annotated

from pydantic import Field

WorkbookPath = Annotated[
    str,
    Field(
        description="Workbook file (.xlsx, .xlsm, .xltx, .xltm): relative to the server's "
        "workbook directory if one is set, else absolute."
    ),
]
SheetName = Annotated[str, Field(description="Worksheet name.")]
CellRef = Annotated[str, Field(description="A single cell in A1 notation, e.g. 'B2'.")]
RangeRef = Annotated[
    str, Field(description="A cell or rectangular range in A1 notation, e.g. 'A1:D20'.")
]
LineIndex = Annotated[
    int, Field(ge=1, description="1-based row number, or 1-based column number (A=1).")
]
LineCount = Annotated[int, Field(ge=1, le=100_000, description="How many rows or columns.")]
