"""Workbooks with content that openpyxl cannot write, assembled by hand for tests."""

import re
import zipfile
from pathlib import Path

from openpyxl import Workbook
from openpyxl.comments import Comment

THREADS = "http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments"
MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
PERSON = "{00000000-0000-4000-8000-000000000001}"
ROOT = "{00000000-0000-4000-8000-0000000000A1}"
REPLY = "{00000000-0000-4000-8000-0000000000A2}"
SECOND = "{00000000-0000-4000-8000-0000000000B1}"
LEGACY = (
    "[Threaded comment]\n\nYour version of Excel allows you to read this threaded comment; "
    "however, any edits to it will get removed if the file is opened in a newer version of "
    "Excel. Learn more: https://go.microsoft.com/fwlink/?linkid=870924\n\nComment:\n    {}"
)
XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'


def _rewrite(
    source: Path,
    target: Path,
    extra: dict[str, str],
    edits: dict[str, tuple[str, str]],
) -> None:
    """Copy a workbook, replacing ``old`` by ``new`` in the parts named in ``edits``."""
    with zipfile.ZipFile(source) as archive, zipfile.ZipFile(target, "w") as out:
        for name in archive.namelist():
            data = archive.read(name).decode("utf-8")
            if name in edits:
                old, new = edits[name]
                data = re.sub(old, new, data)
            out.writestr(name, data)
        for name, content in extra.items():
            out.writestr(name, content)


def threaded_workbook(path: Path) -> None:
    """A workbook with two threaded comments (one with a reply) and a plain note.

    Written by hand after [MS-XLSX]; Excel reads and saves the same structure.
    """
    workbook = Workbook()
    sheet = workbook.active
    assert sheet is not None
    sheet.title = "Data"
    sheet["B2"].comment = Comment(LEGACY.format("First") + "\nReply:\n    Answer", f"tc={ROOT}")
    sheet["C3"].comment = Comment(LEGACY.format("Second"), f"tc={SECOND}")
    sheet["A5"].comment = Comment("A plain note", "Someone")
    plain = path.with_name("plain.xlsx")
    workbook.save(plain)
    stamp = f'dT="2026-10-08T10:00:00.00" personId="{PERSON}"'
    threads = (
        f'{XML}<ThreadedComments xmlns="{THREADS}">'
        f'<threadedComment ref="B2" {stamp} id="{ROOT}"><text>First</text></threadedComment>'
        f'<threadedComment ref="B2" {stamp} id="{REPLY}" parentId="{ROOT}"><text>Answer</text>'
        "</threadedComment>"
        f'<threadedComment ref="C3" {stamp} id="{SECOND}"><text>Second</text></threadedComment>'
        "</ThreadedComments>"
    )
    persons = (
        f'{XML}<personList xmlns="{THREADS}"><person displayName="Test User" id="{PERSON}" '
        'userId="test@example.com" providerId="None"/></personList>'
    )
    ct = "application/vnd.ms-excel"
    rel = "http://schemas.microsoft.com/office/2017/10/relationships"
    _rewrite(
        plain,
        path,
        {"xl/threadedComments/threadedComment1.xml": threads, "xl/persons/person.xml": persons},
        {
            "[Content_Types].xml": (
                "</Types>",
                f'<Override PartName="/xl/threadedComments/threadedComment1.xml" '
                f'ContentType="{ct}.threadedcomments+xml"/>'
                f'<Override PartName="/xl/persons/person.xml" ContentType="{ct}.person+xml"/>'
                "</Types>",
            ),
            "xl/worksheets/_rels/sheet1.xml.rels": (
                "</Relationships>",
                f'<Relationship Id="rIdTC" Type="{rel}/threadedComment" '
                'Target="/xl/threadedComments/threadedComment1.xml"/></Relationships>',
            ),
            "xl/_rels/workbook.xml.rels": (
                "</Relationships>",
                f'<Relationship Id="rIdPerson" Type="{rel}/person" '
                'Target="/xl/persons/person.xml"/></Relationships>',
            ),
        },
    )
    plain.unlink()


def value_metadata_workbook(path: Path) -> None:
    """A workbook whose cells A1 and B1 carry ``vm="1"``, as linked data types do."""
    workbook = Workbook()
    sheet = workbook.active
    assert sheet is not None
    sheet.title = "Sheet"
    sheet["A1"] = "#VALUE!"
    sheet["B1"] = "#VALUE!"
    plain = path.with_name("plain.xlsx")
    workbook.save(plain)
    metadata = (
        f'{XML}<metadata xmlns="{MAIN}" '
        'xmlns:xlrd="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">'
        '<metadataTypes count="1"><metadataType name="XLRICHVALUE" minSupportedVersion="120000"/>'
        '</metadataTypes><futureMetadata name="XLRICHVALUE" count="1"><bk><extLst>'
        '<ext uri="{3e2802c4-a4d2-4d8b-9148-e3be6c30e623}"><xlrd:rvb i="0"/></ext></extLst></bk>'
        '</futureMetadata><valueMetadata count="1"><bk><rc t="1" v="0"/></bk></valueMetadata>'
        "</metadata>"
    )
    sheet_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml"
    sheet_rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata"
    _rewrite(
        plain,
        path,
        {"xl/metadata.xml": metadata},
        {
            "xl/worksheets/sheet1.xml": (r'<c r="([AB]1)"', r'<c r="\1" vm="1"'),
            "[Content_Types].xml": (
                "</Types>",
                f'<Override PartName="/xl/metadata.xml" ContentType="{sheet_type}"/></Types>',
            ),
            "xl/_rels/workbook.xml.rels": (
                "</Relationships>",
                f'<Relationship Id="rIdMeta" Type="{sheet_rel}" Target="/xl/metadata.xml"/>'
                "</Relationships>",
            ),
        },
    )
    plain.unlink()
