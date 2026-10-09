"""The workbook's Company property, which openpyxl cannot hold.

openpyxl rewrites docProps/app.xml on every save without it, so the company is read from the
original file and put into the file that is saved.
"""

import zipfile
from pathlib import Path
from xml.sax.saxutils import escape

from openpyxl import Workbook
from openpyxl.xml.functions import fromstring

APP_PART = "docProps/app.xml"
CORE_PART = "docProps/core.xml"
_DC = "http://purl.org/dc/elements/1.1/"
_NAMESPACE = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"


def read_company(path: Path) -> str:
    with zipfile.ZipFile(path) as archive:
        return company_in(archive)


def forget_default_creator(workbook: Workbook, path: Path) -> None:
    """Leave the creator empty for a file that names none.

    openpyxl fills in its own name for a missing creator, and would save it into the file.
    """
    if workbook.properties.creator != "openpyxl":
        return
    with zipfile.ZipFile(path) as archive:
        core = fromstring(archive.read(CORE_PART)) if CORE_PART in archive.namelist() else None
    if core is not None and core.findtext(f"{{{_DC}}}creator") == "openpyxl":
        return
    workbook.properties.creator = None


def company_in(archive: zipfile.ZipFile) -> str:
    if APP_PART not in archive.namelist():
        return ""
    return fromstring(archive.read(APP_PART)).findtext(f"{{{_NAMESPACE}}}Company", default="")


def named(app_xml: bytes, company: str) -> bytes:
    """openpyxl's app.xml has no company; add one at the end of the root element."""
    closing = b"</Properties>"
    element = b"<Company>" + escape(company).encode() + b"</Company>"
    return app_xml.replace(closing, element + closing)
