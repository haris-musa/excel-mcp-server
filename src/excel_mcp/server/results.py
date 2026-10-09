"""The result shape of tools that change a workbook."""

from pydantic import BaseModel


# What a later call needs to refer to the change. Fields that do not apply are omitted. The
# class has no docstring because pydantic would add it to the schema of every tool.
class Changed(BaseModel):
    path: str | None = None
    sheet: str | None = None
    range: str | None = None
    name: str | None = None
    note: str | None = None
