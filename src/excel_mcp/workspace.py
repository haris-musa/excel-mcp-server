"""Opening and saving workbooks safely."""

import io
import os
import tempfile
import threading
import zipfile
from collections.abc import Iterator
from contextlib import contextmanager
from pathlib import Path
from typing import cast

from openpyxl import Workbook, load_workbook
from openpyxl.chart import AreaChart
from openpyxl.worksheet.worksheet import Worksheet

from excel_mcp.config import Limits
from excel_mcp.errors import (
    InvalidArgumentError,
    LimitExceededError,
    SheetNotFoundError,
    WorkbookError,
    WorkbookExistsError,
    WorkbookNotFoundError,
)
from excel_mcp.paths import MACRO_SUFFIXES, TEMPLATE_SUFFIXES, PathPolicy


class Workspace:
    """Resolves client paths and gives access to workbooks within the limits.

    Changes are serialized: tools run in parallel threads, and each edit loads
    and saves the whole file, so concurrent edits would otherwise lose data.
    """

    def __init__(self, paths: PathPolicy, limits: Limits) -> None:
        self.paths = paths
        self.limits = limits
        self._write_lock = threading.Lock()

    def resolve(self, raw_path: str) -> Path:
        return self.paths.resolve(raw_path)

    def resolve_existing(self, raw_path: str) -> Path:
        path = self.resolve(raw_path)
        if not path.is_file():
            raise WorkbookNotFoundError(f"Workbook {self.display(path)} does not exist.")
        size = path.stat().st_size
        if size > self.limits.max_file_bytes:
            raise LimitExceededError(
                f"Workbook is {size:,} bytes; the limit is {self.limits.max_file_bytes:,}."
            )
        return path

    def resolve_directory(self, raw_path: str) -> Path:
        return self.paths.resolve_directory(raw_path)

    def display(self, path: Path) -> str:
        return self.paths.display(path)

    @contextmanager
    def read(self, raw_path: str, *, data_only: bool = False) -> Iterator[Workbook]:
        """Open a workbook for reading; changes are never saved."""
        workbook = self._load(self.resolve_existing(raw_path), data_only=data_only)
        try:
            yield workbook
        finally:
            close_workbook(workbook)

    @contextmanager
    def edit(self, raw_path: str) -> Iterator[Workbook]:
        """Open a workbook and save it only if the block finishes without error."""
        with self._write_lock:
            path = self.resolve_existing(raw_path)
            workbook = self._load(path, data_only=False)
            _restore_area_chart_axes(workbook)
            try:
                yield workbook
                save_atomically(workbook, path)
            finally:
                close_workbook(workbook)

    def create(self, raw_path: str, sheets: list[str], *, overwrite: bool) -> Path:
        path = self.resolve(raw_path)
        if path.suffix.lower() in MACRO_SUFFIXES:
            raise InvalidArgumentError(
                "New workbooks cannot contain macros. Create an .xlsx or .xltx file instead."
            )
        workbook = Workbook()
        workbook.worksheets[0].title = sheets[0]
        for name in sheets[1:]:
            workbook.create_sheet(name)
        workbook.template = path.suffix.lower() in TEMPLATE_SUFFIXES
        with self._write_lock:
            self._check_overwrite(path, overwrite)
            save_atomically(workbook, path)
        return path

    def store(self, raw_path: str, content: bytes, *, overwrite: bool) -> Path:
        """Save uploaded workbook bytes as a file."""
        path = self.resolve(raw_path)
        with self._write_lock:
            self._check_overwrite(path, overwrite)
            write_atomically(path, content)
        return path

    def _check_overwrite(self, path: Path, overwrite: bool) -> None:
        if path.exists() and not overwrite:
            raise WorkbookExistsError(
                f"{self.display(path)} already exists. Pass overwrite=true to replace it."
            )

    def _load(self, path: Path, *, data_only: bool) -> Workbook:
        if not zipfile.is_zipfile(path):
            raise WorkbookError(
                f"{self.display(path)} is not an Excel workbook. Only .xlsx/.xlsm files are "
                "supported; legacy .xls and CSV files must be converted first."
            )
        try:
            return load_workbook(
                path, data_only=data_only, keep_vba=path.suffix.lower() in MACRO_SUFFIXES
            )
        except Exception as error:
            # openpyxl raises many different exception types for damaged files.
            raise WorkbookError(
                f"{self.display(path)} could not be opened as an Excel workbook ({error})."
            ) from None


def save_atomically(workbook: Workbook, path: Path) -> None:
    """Serialize the workbook in memory, then write it atomically."""
    buffer = io.BytesIO()
    try:
        workbook.save(buffer)
    except Exception as error:
        # openpyxl reports invalid workbook states with many exception types.
        raise WorkbookError(f"Could not save the workbook: {error}") from None
    write_atomically(path, buffer.getvalue())


def write_atomically(path: Path, content: bytes) -> None:
    """Write to a new temporary file next to the target, then swap it in.

    mkstemp creates the file exclusively, so a symlink planted at the temporary
    name cannot redirect the write, and a failed write leaves the target intact.
    """
    path.parent.mkdir(parents=True, exist_ok=True)
    handle, temp_name = tempfile.mkstemp(dir=path.parent, prefix=".~", suffix=".tmp")
    temp_path = Path(temp_name)
    try:
        with os.fdopen(handle, "wb") as file:
            file.write(content)
        temp_path.replace(path)
    except OSError as error:
        raise WorkbookError(f"Could not save the workbook: {error.strerror}.") from None
    finally:
        temp_path.unlink(missing_ok=True)


def close_workbook(workbook: Workbook) -> None:
    workbook.close()
    # openpyxl keeps macro parts in an in-memory zip that it never closes.
    if workbook.vba_archive:
        workbook.vba_archive.close()


def _restore_area_chart_axes(workbook: Workbook) -> None:
    """Undo an openpyxl bug that hides the axes of area charts.

    openpyxl drops the axes' ``delete`` flag when it reads an area chart, and
    Excel treats the missing flag as hidden, so without this every edit would
    remove the axes of area charts already in the workbook.
    """
    for sheet in worksheets(workbook):
        for chart in sheet._charts:  # pyright: ignore[reportAttributeAccessIssue]
            if isinstance(chart, AreaChart):
                for axis in (chart.x_axis, chart.y_axis):
                    if axis.delete is None:
                        axis.delete = False


def get_sheet(workbook: Workbook, name: str) -> Worksheet:
    if name not in workbook.sheetnames:
        raise SheetNotFoundError(name, workbook.sheetnames)
    sheet = workbook[name]
    if not isinstance(sheet, Worksheet):
        raise SheetNotFoundError(name, [sheet.title for sheet in worksheets(workbook)])
    return sheet


def worksheets(workbook: Workbook) -> list[Worksheet]:
    """The workbook's worksheets, leaving out chart sheets."""
    return [sheet for sheet in workbook.worksheets if isinstance(sheet, Worksheet)]


def sheet_names(sheet: Worksheet) -> list[str]:
    """The names of all sheets in the workbook that holds ``sheet``."""
    # openpyxl types parent as optional, but every sheet it loads or creates has one.
    return cast(Workbook, sheet.parent).sheetnames
