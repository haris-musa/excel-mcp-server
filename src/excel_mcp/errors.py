"""Errors raised by the domain layer.

The server turns every `ExcelMCPError` into an MCP tool error, and its message
is shown to the model, so messages should say what went wrong and how to fix it.
"""


class ExcelMCPError(Exception):
    """Base class for expected, user-facing errors."""


class InvalidArgumentError(ExcelMCPError):
    """A tool argument is malformed."""


class PathNotAllowedError(ExcelMCPError):
    """A path is outside the allowed directories or has a disallowed type."""


class WorkbookNotFoundError(ExcelMCPError):
    """The workbook file does not exist."""


class WorkbookExistsError(ExcelMCPError):
    """The workbook file already exists."""


class WorkbookError(ExcelMCPError):
    """The workbook could not be opened or saved."""


class SheetNotFoundError(ExcelMCPError):
    """The worksheet does not exist."""

    def __init__(self, sheet: str, available: list[str]) -> None:
        names = ", ".join(repr(name) for name in available)
        super().__init__(f"Sheet {sheet!r} not found. Available sheets: {names}.")


class UnsafeFormulaError(ExcelMCPError):
    """A formula was rejected by the formula safety policy."""


class LimitExceededError(ExcelMCPError):
    """A request exceeds a configured size limit."""
