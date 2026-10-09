"""What describe_sheet says about a chart."""

from pydantic import BaseModel


class ChartInfo(BaseModel):
    name: str
    type: str
    title: str | None = None
    range: str | None = None
    series: list[str] = []
