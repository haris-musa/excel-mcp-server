"""Sheet protection with an optional password."""

from base64 import b64decode
from hashlib import sha512
from typing import Literal

from openpyxl.utils.protection import hash_password
from openpyxl.worksheet.protection import SheetProtection
from openpyxl.worksheet.worksheet import Worksheet
from pydantic import Field

from excel_mcp.errors import InvalidArgumentError
from excel_mcp.inputs import InputModel

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


class Protection(InputModel):
    enabled: bool = Field(description="False unprotects.")
    password: str | None = Field(
        default=None, max_length=255, description="To set; to unprotect, the current one."
    )
    allow: list[AllowedAction] = Field(
        default=["select_locked_cells", "select_unlocked_cells"],
        description="What users may still do.",
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
    return password_matches(
        password,
        legacy=current.password,
        algorithm=current.algorithmName,
        salt=current.saltValue,
        spin_count=current.spinCount,
        digest=current.hashValue,
    )


def password_matches(
    password: str | None,
    *,
    legacy: str | None,
    algorithm: str | None,
    salt: str | None,
    spin_count: int | None,
    digest: str | None,
) -> bool:
    """Check a password against what Excel stored: the old 16-bit hash or a salted SHA-512."""
    if password is None:
        return False
    if not digest:
        return int(hash_password(password), 16) == int(legacy or "0", 16)
    if algorithm != "SHA-512":
        raise InvalidArgumentError(
            f"The password uses {algorithm}, which is not supported. Unprotect it in Excel."
        )
    hashed = sha512(b64decode(salt or "") + password.encode("utf-16-le")).digest()
    for index in range(spin_count or 0):
        hashed = sha512(hashed + index.to_bytes(4, "little")).digest()
    return hashed == b64decode(digest)


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
