"""The digits that Excel reads as 0-9 besides the ASCII ones.

Excel 16 reads the native digits of these scripts like ASCII digits wherever text becomes a
number, a date or a reference (typing, VALUE, criteria, INDIRECT). Any other Unicode digit,
such as a fullwidth one, is a plain character to it.
"""

_ZEROS = (
    0x0660,  # Arabic-Indic
    0x06F0,  # Extended Arabic-Indic
    0x0966,  # Devanagari
    0x09E6,  # Bengali
    0x0A66,  # Gurmukhi
    0x0AE6,  # Gujarati
    0x0B66,  # Oriya
    0x0BE6,  # Tamil
    0x0C66,  # Telugu
    0x0CE6,  # Kannada
    0x0D66,  # Malayalam
    0x0E50,  # Thai
    0x0ED0,  # Lao
    0x0F20,  # Tibetan
    0x1040,  # Myanmar
    0x17E0,  # Khmer
    0x1810,  # Mongolian
)
_TO_ASCII = {zero + digit: str(digit) for zero in _ZEROS for digit in range(10)}


def ascii_digits(text: str) -> str:
    """``text`` with the digits Excel reads as 0-9 written as ASCII digits."""
    return text.translate(_TO_ASCII)
