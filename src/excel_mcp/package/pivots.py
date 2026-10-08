"""The data fields and page fields of a PivotTable definition, where Excel keeps extensions."""

import re

FIELD_KINDS = ("dataField", "pageField")


def fields(xml: str, kind: str) -> list[str]:
    """The text of each ``<dataField>`` or ``<pageField>`` element, in order."""
    return [m[0] for m in _pattern(kind).finditer(xml)]


def with_field_extensions(xml: str, kind: str, extensions: dict[int, str]) -> str:
    """The definition with an ``<extLst>`` put into the fields at these positions.

    openpyxl writes the fields as empty elements, which the extension list turns into
    elements with content.
    """

    pieces: list[str] = []
    last = 0
    for position, match in enumerate(_pattern(kind).finditer(xml)):
        extension = extensions.get(position)
        if extension is not None and match[0].endswith("/>"):
            pieces += [xml[last : match.start()], f"{match[0][:-2]}>{extension}</{kind}>"]
            last = match.end()
    return "".join(pieces) + xml[last:]


def _pattern(kind: str) -> re.Pattern[str]:
    return re.compile(rf"<{kind}\b[^>]*?(?:/>|>.*?</{kind}>)", re.S)
