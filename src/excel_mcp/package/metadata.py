"""``xl/metadata.xml``: where Excel keeps what a cell's ``cm`` attribute points to.

A dynamic array formula carries ``cm="N"``, the Nth ``<bk>`` of ``cellMetadata``, which says
"this array may spill": a ``XLDAPR`` metadata type whose future metadata sets ``fDynamic``.
The same part also holds the value metadata (``vm``) of linked data types and images in
cells, so adding to it must keep what is there.
"""

import re

from excel_mcp.errors import WorkbookError
from excel_mcp.package.scan import scan

PART = "xl/metadata.xml"
CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml"
RELATIONSHIP = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata"
_XDA = "http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray"
_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"

_ORDER = ["metadataTypes", "futureMetadata", "cellMetadata", "valueMetadata", "extLst"]
_TYPE = (
    '<metadataType name="XLDAPR" minSupportedVersion="120000" copy="1" pasteAll="1" '
    'pasteValues="1" merge="1" splitFirst="1" rowColShift="1" clearFormats="1" '
    'clearComments="1" assign="1" coerce="1" cellMeta="1"/>'
)
_PROPERTIES = (
    '<bk><extLst><ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}">'
    '<xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/></ext></extLst></bk>'
)
EMPTY = (
    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<metadata xmlns="{_MAIN}" '
    f'xmlns:xda="{_XDA}"><metadataTypes count="0"></metadataTypes></metadata>'
)


class Metadata:
    """The text of ``xl/metadata.xml``, extended in place."""

    def __init__(self, xml: str | None) -> None:
        self.xml = xml or EMPTY
        if "<metadata" not in self.xml:
            raise WorkbookError("xl/metadata.xml has an unexpected layout.")
        if "xmlns:xda=" not in self.xml:
            self.xml = self.xml.replace("<metadata ", f'<metadata xmlns:xda="{_XDA}" ', 1)

    def dynamic_array_index(self) -> int:
        """The ``cm`` value of dynamic array formulas, added to the metadata if it is missing."""
        type_index = self._index("metadataTypes", "metadataType", 'name="XLDAPR"', _TYPE)
        properties = self._index(
            "futureMetadata",
            "bk",
            'fDynamic="1" fCollapsed="0"',
            _PROPERTIES,
            block='name="XLDAPR"',
            attributes=' name="XLDAPR"',
        )
        reference = f'<rc t="{type_index}" v="{properties - 1}"/>'
        return self._index("cellMetadata", "bk", reference, f"<bk>{reference}</bk>")

    def _index(
        self,
        section: str,
        item: str,
        marker: str,
        new: str,
        block: str = "",
        attributes: str = "",
    ) -> int:
        """The 1-based position of the item holding ``marker``; added to its section if absent."""
        opening = rf"<{section}\b(?=[^>]*{block})[^>]*>" if block else rf"<{section}\b[^>]*>"
        found = re.search(rf"({opening})(.*?)(</{section}>)", self.xml, re.S)
        if found is None:
            found = self._create(section, attributes)
        items = re.findall(rf"<{item}\b(?:[^>]*?/>|.*?</{item}>)", found[2], re.S)
        for position, text in enumerate(items, 1):
            if marker in text:
                return position
        body = found[2] + new
        opening_tag = re.sub(r'count="\d+"', f'count="{len(items) + 1}"', found[1])
        self.xml = (
            self.xml[: found.start()] + opening_tag + body + found[3] + self.xml[found.end() :]
        )
        return len(items) + 1

    def _create(self, section: str, attributes: str) -> re.Match[str]:
        """Add an empty section where the schema puts it: after the ones that precede it."""
        rank = {name: position for position, name in enumerate(_ORDER)}
        data = self.xml.encode("utf-8")
        later = (c for c in scan(data).children if rank.get(c.local, 0) > rank[section])
        following = next(later, None)
        point = following.start if following else scan(data).close_tag
        empty = f'<{section}{attributes} count="0"></{section}>'.encode()
        self.xml = (data[:point] + empty + data[point:]).decode("utf-8")
        found = re.search(rf"(<{section}\b[^>]*>)()(</{section}>)", self.xml)
        if found is None:
            raise WorkbookError("xl/metadata.xml has an unexpected layout.")
        return found
