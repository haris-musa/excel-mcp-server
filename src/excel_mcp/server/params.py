"""Parameter types shared by the tools, with the descriptions clients show the model."""

from typing import Annotated

from pydantic import Field

WorkbookPath = Annotated[
    str,
    Field(description="Workbook path: relative to the server's workbook folder, or absolute."),
]
SheetName = Annotated[str, Field(description="Worksheet name.")]
CellRef = Annotated[str, Field(description="Cell, e.g. 'B2'.")]
RangeRef = Annotated[str, Field(description="Cell or range, e.g. 'A1:D20'.")]
LineIndex = Annotated[
    int, Field(ge=1, description="1-based row number, or 1-based column number (A=1).")
]
LineCount = Annotated[int, Field(ge=1, le=100_000, description="How many.")]
