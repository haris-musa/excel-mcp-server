"""Sheet protection with an optional password."""

from base64 import b64decode
from hashlib import sha512
from typing import Literal

from openpyxl.utils.protection import hash_password
from openpyxl.worksheet.protection import SheetProtection
from openpyxl.worksheet.worksheet import Worksheet
from pydantic import BaseModel, Field

from excel_mcp.errors import InvalidArgumentError

AllowedAction = Literal[
    "select_locked_cells",
    "select_unlocked_cells",
    "format_cells",
    "format_columns",
    "format_rows",
    "insert_columns",
    "insert_rows",
    "insert_hyperlinks",
    "delete_columns",
    "delete_rows",
    "sort",
    "auto_filter",
    "pivot_tables",
]
ALLOWED_ACTIONS = AllowedAction.__args__


class Protection(BaseModel):
    """Turn sheet protection on or off."""

    enabled: bool = Field(description="True protects the sheet, false unprotects it.")
    password: str | None = Field(
        default=None,
        max_length=255,
        description="Password to set; when unprotecting, the password the sheet has.",
    )
    allow: list[AllowedAction] = Field(
        default=["select_locked_cells", "select_unlocked_cells"],
        description="Actions users may still do on a protected sheet.",
    )


def apply_protection(sheet: Worksheet, change: Protection) -> None:
    current = sheet.protection
    secured = bool(current.sheet and (current.password or current.hashValue))
    if change.enabled:
        if secured:
            raise InvalidArgumentError(
                "The sheet is already protected with a password. Unprotect it first."
            )
        sheet.protection = _protected(change)
    else:
        if secured and not _password_matches(current, change.password):
            raise InvalidArgumentError("The sheet is password protected: give the right password.")
        sheet.protection = SheetProtection()


def _password_matches(current: SheetProtection, password: str | None) -> bool:
    if password is None:
        return False
    if current.hashValue:
        return _sha512_hash(current, password) == b64decode(current.hashValue)
    return int(hash_password(password), 16) == int(current.password, 16)


def _sha512_hash(current: SheetProtection, password: str) -> bytes:
    """The hash Excel 2013 and later store: salted, then repeated spinCount times."""
    if current.algorithmName != "SHA-512":
        raise InvalidArgumentError(
            f"This sheet's password uses {current.algorithmName}, which is not supported. "
            "Unprotect it in Excel."
        )
    digest = sha512(b64decode(current.saltValue) + password.encode("utf-16-le")).digest()
    for index in range(current.spinCount or 0):
        digest = sha512(digest + index.to_bytes(4, "little")).digest()
    return digest


def _protected(change: Protection) -> SheetProtection:
    protection = SheetProtection(sheet=True, objects=True, scenarios=True)
    for action in ALLOWED_ACTIONS:
        setattr(protection, _camel_case(action), action not in change.allow)
    if change.password:
        protection.set_password(change.password)
    return protection


def _camel_case(name: str) -> str:
    first, *rest = name.split("_")
    return first + "".join(word.capitalize() for word in rest)
