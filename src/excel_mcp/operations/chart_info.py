"""What describe_sheet says about a chart."""

from pydantic import BaseModel


class ChartInfo(BaseModel):
    index: int
    type: str
    title: str | None
    anchor: str | None
    series: list[str]
