"""The CSV dialect options shared by import and export."""

from typing import Literal

Delimiter = Literal["comma", "semicolon", "tab", "pipe"]
Encoding = Literal["utf-8", "utf-8-sig", "cp1252", "utf-16"]

DELIMITER_CHARACTERS: dict[str, str] = {"comma": ",", "semicolon": ";", "tab": "\t", "pipe": "|"}
FORMULA_STARTS = ("=", "+", "-", "@", "\t", "\r")  # text a spreadsheet would run as a formula
