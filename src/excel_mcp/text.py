"""Formatting shared by every message that lists names for the model."""

from collections.abc import Iterable


def quoted(items: Iterable[object], conjunction: str | None = None) -> str:
    """``'a', 'b'``, or ``'a', 'b' or 'c'`` with ``conjunction="or"``."""
    names = [repr(str(item)) for item in items]
    if conjunction and len(names) > 1:
        return f"{', '.join(names[:-1])} {conjunction} {names[-1]}"
    return ", ".join(names)
