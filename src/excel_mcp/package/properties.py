"""The workbook's Company property, which openpyxl cannot hold.

openpyxl rewrites docProps/app.xml on every save without it, so the company is read from the
original file and put into the file that is saved.
"""

import zipfile
from pathlib import Path
from xml.sax.saxutils import escape

from openpyxl.xml.functions import fromstring

APP_PART = "docProps/app.xml"
_NAMESPACE = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"


def read_company(path: Path) -> str:
    with zipfile.ZipFile(path) as archive:
        return company_in(archive)


def company_in(archive: zipfile.ZipFile) -> str:
    if APP_PART not in archive.namelist():
        return ""
    return fromstring(archive.read(APP_PART)).findtext(f"{{{_NAMESPACE}}}Company", default="")


def named(app_xml: bytes, company: str) -> bytes:
    """openpyxl's app.xml has no company; add one at the end of the root element."""
    closing = b"</Properties>"
    element = b"<Company>" + escape(company).encode() + b"</Company>"
    return app_xml.replace(closing, element + closing)
