"""The workbook's Company property.

openpyxl rewrites docProps/app.xml on every save and cannot hold a company, so the company
is read when a workbook is opened for editing and written back into the saved file.
"""

import io
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape

from openpyxl.workbook import Workbook
from openpyxl.xml.functions import fromstring

APP_PART = "docProps/app.xml"
NAMESPACE = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
_COMPANY = f"{{{NAMESPACE}}}Company"


def read_company(path: Path) -> str:
    with zipfile.ZipFile(path) as archive:
        if APP_PART not in archive.namelist():
            return ""
        return fromstring(archive.read(APP_PART)).findtext(_COMPANY, default="")


def attach_company(workbook: Workbook, company: str) -> None:
    workbook.__dict__["company"] = company


def company_of(workbook: Workbook) -> str:
    return workbook.__dict__.get("company", "")


def with_company(content: bytes, company: str) -> bytes:
    """The saved workbook ``content`` with its app.xml naming ``company``."""
    if not company:
        return content
    output = io.BytesIO()
    with (
        zipfile.ZipFile(io.BytesIO(content)) as source,
        zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as target,
    ):
        for entry in source.infolist():
            data = source.read(entry.filename)
            if entry.filename == APP_PART:
                data = _named(data, company)
            target.writestr(entry, data)
    return output.getvalue()


def _named(app_xml: bytes, company: str) -> bytes:
    """openpyxl's app.xml has no company; add one at the end of the root element."""
    closing = b"</Properties>"
    return app_xml.replace(
        closing, b"<Company>" + escape(company).encode() + b"</Company>" + closing
    )
