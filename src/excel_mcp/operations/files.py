"""Moving whole workbook files in and out of the server as base64."""

import base64
import binascii
import io
import zipfile
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.worksheet.formula import ArrayFormula

from excel_mcp.errors import InvalidArgumentError, LimitExceededError, UnsafeFormulaError
from excel_mcp.formulas import check_formula
from excel_mcp.workspace import close_workbook, worksheets


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


def _check_formulas(content: bytes) -> None:
    try:
        workbook = load_workbook(io.BytesIO(content))
    except Exception as error:
        # openpyxl raises many different exception types for damaged files.
        raise InvalidArgumentError(f"The uploaded workbook could not be read ({error}).") from None
    try:
        if workbook._external_links:  # pyright: ignore[reportAttributeAccessIssue]
            raise UnsafeFormulaError("Workbooks that link to other workbooks cannot be uploaded.")
        for sheet in worksheets(workbook):
            for cell in sheet._cells.values():
                if cell.data_type == "f":
                    formula = (
                        cell.value.text if isinstance(cell.value, ArrayFormula) else cell.value
                    )
                    check_formula(str(formula))
    finally:
        close_workbook(workbook)
