"""Which external addresses a hyperlink may have, for links the server writes and for uploads."""

from urllib.parse import urlsplit

SCHEMES = frozenset({"http", "https", "mailto"})


def is_allowed_address(target: str) -> bool:
    """A full http, https or mailto address; web addresses may not carry credentials.

    Everything else opens something outside the browser or the mail client when clicked:
    file paths, network shares, ``smb:``, ``ms-excel:``, ``search-ms:``, ``javascript:``.
    """
    if any(ord(char) < 32 for char in target):
        return False
    parts = urlsplit(target)
    if parts.scheme not in SCHEMES:
        return False
    if parts.scheme == "mailto":
        return bool(parts.path)
    return bool(parts.netloc) and "@" not in parts.netloc and "\\" not in parts.netloc
