"""Moving whole workbook files in and out of the server as base64."""

import base64
import binascii
import io
import zipfile
from pathlib import Path
from typing import NamedTuple

from openpyxl import load_workbook

from excel_mcp.config import Limits
from excel_mcp.errors import (
    InvalidArgumentError,
    InvalidFormulaError,
    LimitExceededError,
    UnsafeFormulaError,
)
from excel_mcp.formulas import check_formula
from excel_mcp.upload_links import RemovedLink, neutralise_links
from excel_mcp.upload_scan import scan_package
from excel_mcp.workspace import close_workbook


def encode_file(path: Path) -> str:
    return base64.b64encode(path.read_bytes()).decode("ascii")


class Upload(NamedTuple):
    content: bytes
    # What was taken out of the file because it reaches outside it; empty for most workbooks.
    removed: list[RemovedLink]


def decode_workbook(content_base64: str, limits: Limits) -> Upload:
    """Decode an uploaded workbook and check it the way the server checks its own writes.

    Hyperlinks to files and network locations are removed; what else reaches outside the
    file is refused.
    """
    max_bytes = limits.max_file_bytes
    if len(content_base64) > max_bytes * 4 // 3 + 4:
        raise LimitExceededError(f"File is larger than the limit of {max_bytes:,} bytes.")
    try:
        content = base64.b64decode(content_base64, validate=True)
    except binascii.Error:
        raise InvalidArgumentError("content_base64 is not valid base64.") from None
    if len(content) > max_bytes:
        raise LimitExceededError(f"File is {len(content):,} bytes; the limit is {max_bytes:,}.")
    if not zipfile.is_zipfile(io.BytesIO(content)):
        raise InvalidArgumentError("The uploaded content is not an .xlsx/.xlsm workbook.")
    # The first pass validates the whole package, so it is only rewritten once it is safe.
    for _ in scan_package(content, limits, check_links=False):
        pass
    content, removed = neutralise_links(content)
    _check_formulas(content, limits)
    return Upload(content, removed)


# The uploaded bytes are stored unchanged, so the check reads the XML itself
# rather than what openpyxl loads: openpyxl drops reserved names such as print
# areas and whole features (x14 rules, sparklines) that would then go unchecked.
def _check_formulas(content: bytes, limits: Limits) -> None:
    sheet_names = _sheet_names(content)
    for formula in scan_package(content, limits):
        try:
            check_formula(f"={formula.text}", sheet_names)
        except InvalidFormulaError:
            if not formula.pivot:
                raise


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
