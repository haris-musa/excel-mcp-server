"""Client text inside the XML the server writes itself."""

import re
from xml.sax.saxutils import escape

# What XML 1.0 cannot hold: control characters other than tab, line feed and carriage return,
# lone surrogates and the two non-characters.
ILLEGAL_CHARACTERS = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f\ud800-\udfff￾￿]")
_ATTRIBUTE_ESCAPES = {'"': "&quot;", "\n": "&#10;", "\r": "&#13;", "\t": "&#9;"}


def attribute(value: str) -> str:
    """The value as a double-quoted attribute, such as ``"a &amp; b"``.

    Unlike ``quoteattr``, it never switches to single quotes, which the regular expressions
    that read these parts back do not expect.
    """
    return f'"{escape(value, _ATTRIBUTE_ESCAPES)}"'
