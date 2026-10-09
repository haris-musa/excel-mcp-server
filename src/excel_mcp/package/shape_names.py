"""The names of charts and pictures.

openpyxl forgets the name Excel gave a chart ("Chart 1") when it reads a drawing. The package
layer reads the names from the file and keeps them on the objects openpyxl made, so that they
are written back and tools can select a shape by its name.
"""

import re
from html import unescape

from openpyxl.chart._chart import ChartBase
from openpyxl.drawing.image import Image

_KEY = "excel_mcp_name"
_SHAPE_NAME = re.compile(r"<(?:\w+:)?cNvPr\b[^>]*?\bname=\"([^\"]*)\"")


def name_of(shape: ChartBase | Image) -> str | None:
    return vars(shape).get(_KEY)


def set_name(shape: ChartBase | Image, name: str) -> None:
    vars(shape)[_KEY] = name


def anchored_name(xml: str) -> str | None:
    """The name of the first shape of an anchor's XML."""
    found = _SHAPE_NAME.search(xml)
    return unescape(found[1]) if found and found[1] else None


def restore_names(shapes: list[ChartBase] | list[Image], names: list[str | None]) -> None:
    """Name `shapes` by `names`, which come in the order openpyxl lists them, if there are as
    many names as shapes."""
    if len(shapes) == len(names):
        for shape, name in zip(shapes, names, strict=True):
            if name:
                set_name(shape, name)
