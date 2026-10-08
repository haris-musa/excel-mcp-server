"""Moving whole workbook files in and out of the server as base64."""

import base64
import binascii
import io
import zipfile
from collections.abc import Iterator
from pathlib import Path
from xml.etree import ElementTree

from openpyxl import load_workbook

from excel_mcp.errors import InvalidArgumentError, LimitExceededError, UnsafeFormulaError
from excel_mcp.formulas import check_formula
from excel_mcp.workspace import close_workbook


def encode_file(path: Path) -> str:
    return base64.b64encode(path.read_bytes()).decode("ascii")


def decode_workbook(content_base64: str, max_bytes: int) -> bytes:
    """Decode an uploaded workbook and check it the way the server checks its own writes."""
    try:
        content = base64.b64decode(content_base64, validate=True)
    except binascii.Error:
        raise InvalidArgumentError("content_base64 is not valid base64.") from None
    if len(content) > max_bytes:
        raise LimitExceededError(f"File is {len(content):,} bytes; the limit is {max_bytes:,}.")
    if not zipfile.is_zipfile(io.BytesIO(content)):
        raise InvalidArgumentError("The uploaded content is not an .xlsx/.xlsm workbook.")
    _check_formulas(content)
    return content


# The uploaded bytes are stored unchanged, so the check reads the XML itself
# rather than what openpyxl loads: openpyxl drops reserved names such as print
# areas and whole features (x14 rules, sparklines) that would then go unchecked.
# These elements hold formulas: cell formulas and chart or sparkline references
# ("f"), rule formulas, table column formulas and defined names.
_FORMULA_ELEMENTS = frozenset(
    {
        "f",
        "formula",
        "formula1",
        "formula2",
        "calculatedColumnFormula",
        "totalsRowFormula",
        "definedName",
    }
)


def _check_formulas(content: bytes) -> None:
    sheet_names = _sheet_names(content)
    with zipfile.ZipFile(io.BytesIO(content)) as archive:
        for part in archive.namelist():
            if part.startswith("xl/") and part.endswith(".xml"):
                for formula in _formulas(archive.read(part), part):
                    check_formula(f"={formula}", sheet_names)


def _sheet_names(content: bytes) -> list[str]:
    try:
        workbook = load_workbook(io.BytesIO(content), read_only=True)
    except Exception as error:
        # openpyxl raises many different exception types for damaged files.
        raise InvalidArgumentError(f"The uploaded workbook could not be read ({error}).") from None
    try:
        if workbook._external_links:  # pyright: ignore[reportAttributeAccessIssue]
            raise UnsafeFormulaError("Workbooks that link to other workbooks cannot be uploaded.")
        return workbook.sheetnames
    finally:
        close_workbook(workbook)


def _formulas(xml: bytes, part: str) -> Iterator[str]:
    try:
        for _, element in ElementTree.iterparse(io.BytesIO(xml)):
            name = element.tag.rpartition("}")[2]
            if name in _FORMULA_ELEMENTS and element.text and element.text.strip():
                yield element.text
            # Color scale, data bar and icon set thresholds can be formulas too.
            elif name == "cfvo" and (value := element.get("val")):
                yield value
    except ElementTree.ParseError:
        raise InvalidArgumentError(f"The uploaded workbook part {part} is not valid XML.") from None
