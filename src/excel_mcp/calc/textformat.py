"""Number formats for TEXT, DOLLAR and FIXED (US English).

Formats that use features not implemented here (fractions, scientific notation,
conditions, colours, elapsed time) raise `UncalculableError` instead of guessing.
"""

import datetime as dt
import re
from decimal import ROUND_HALF_UP, Decimal

from excel_mcp.calc.values import UncalculableError, number_text, serial_to_date

_DATE_TOKEN = re.compile(r"yyyy|yy|y|mmmm|mmm|mm|m|dddd|ddd|dd|d|hh|h|ss|s|AM/PM|A/P", re.I)
_LITERALS = set("$-+/():!^&'~{}<>= ")
_MONTHS = [dt.date(2000, month, 1).strftime("%B") for month in range(1, 13)]
_DAYS = [dt.date(2000, 1, 3 + day).strftime("%A") for day in range(7)]


def format_number(value: float, fmt: str) -> str:
    if fmt.casefold() == "general":
        return number_text(value)
    sections = _split_sections(fmt)
    if len(sections) > 3:
        raise UncalculableError("text section in number format")
    negative = value < 0
    if len(sections) == 1:
        section = sections[0]
    elif negative:
        section, negative = sections[1], False
    elif value == 0 and len(sections) == 3:
        section = sections[2]
    else:
        section = sections[0]
    if _is_date(section):
        return _format_date(abs(value), section)
    return ("-" if negative else "") + _format_digits(abs(value), section)


def format_fixed(value: float, decimals: int, grouping: bool) -> str:
    """Round half away from zero and write with ``decimals`` places."""
    exact = Decimal(f"{abs(value):.15g}")
    quantum = Decimal(1).scaleb(-decimals)
    rounded = exact.quantize(quantum, rounding=ROUND_HALF_UP) if decimals >= 0 else None
    if rounded is None:
        rounded = (exact.scaleb(decimals).quantize(Decimal(1), rounding=ROUND_HALF_UP)).scaleb(
            -decimals
        )
        decimals = 0
    text = f"{rounded:,.{decimals}f}" if grouping else f"{rounded:.{decimals}f}"
    return ("-" if value < 0 and rounded != 0 else "") + text


def _split_sections(fmt: str) -> list[str]:
    sections, current, quoted = [], "", False
    for character in fmt:
        if character == '"':
            quoted = not quoted
        if character == ";" and not quoted:
            sections.append(current)
            current = ""
        else:
            current += character
    return [*sections, current]


def _is_date(section: str) -> bool:
    return bool(re.search(r"[ydhs]|(?<![0#?])m", _strip_quoted(section), re.I))


def _strip_quoted(section: str) -> str:
    return re.sub(r'"[^"]*"|\\.', "", section)


def _format_digits(value: float, section: str) -> str:
    if re.search(r"[eE][+-]|\[|\?|\*|@|/", _strip_quoted(section)):
        raise UncalculableError("number format")
    prefix, pattern, suffix = _split_pattern(section)
    scale = _strip_quoted(pattern + suffix + prefix).count("%")
    shown = Decimal(f"{value:.15g}") * Decimal(100) ** scale
    trailing = len(pattern) - len(pattern.rstrip(","))
    shown = shown.scaleb(-3 * trailing)
    pattern = pattern.rstrip(",")
    grouping = "," in pattern
    whole_mask, _, fraction_mask = pattern.replace(",", "").partition(".")
    decimals = len(fraction_mask)
    rounded = shown.quantize(Decimal(1).scaleb(-decimals), rounding=ROUND_HALF_UP)
    whole, _, fraction = f"{rounded:f}".partition(".")
    whole = whole.lstrip("0")
    required = whole_mask.count("0")
    whole = whole.rjust(required, "0")
    if grouping and whole:
        whole = f"{int(whole):,}".rjust(len(whole), "0")
    keep = len(fraction_mask.rstrip("#"))
    fraction = fraction[:keep] + fraction[keep:].rstrip("0")
    dot = "." if "." in pattern else ""
    return prefix + whole + dot + fraction + suffix


def _split_pattern(section: str) -> tuple[str, str, str]:
    """Literal text before, the digit placeholders of, and literal text after a number format."""
    prefix, pattern, suffix = "", "", ""
    index = 0
    state = 0
    while index < len(section):
        character = section[index]
        if character == '"':
            end = section.index('"', index + 1)
            literal = section[index + 1 : end]
            index = end
        elif character == "\\":
            index += 1
            literal = section[index]
        elif character == "_":
            index += 1
            literal = " "
        elif character == "%":
            literal = "%"
        elif character in "0#.,":
            if state == 2:
                raise UncalculableError("number format")
            state = 1
            pattern += character
            index += 1
            continue
        elif character in _LITERALS:
            literal = character
        else:
            raise UncalculableError("number format")
        if state == 1:
            state = 2
        if state == 0:
            prefix += literal
        else:
            suffix += literal
        index += 1
    return prefix, pattern, suffix


def _format_date(serial: float, section: str) -> str:
    day = serial_to_date(serial)
    seconds = round((serial % 1) * 86400)
    if seconds >= 86400:
        seconds, day = 0, day + dt.timedelta(days=1)
    hour, minute, second = seconds // 3600, seconds // 60 % 60, seconds % 60
    twelve_hour = bool(re.search(r"AM/PM|A/P", section, re.I))
    parts = re.split(r'("[^"]*"|\\.|\[[^\]]*\])', section)
    tokens: list[tuple[bool, str]] = []
    for part in parts:
        if part.startswith("["):
            raise UncalculableError("elapsed time format")
        if part.startswith('"'):
            tokens.append((False, part[1:-1]))
        elif part.startswith("\\"):
            tokens.append((False, part[1]))
        else:
            tokens.extend(_split_date_tokens(part))
    out = []
    for index, (is_token, text) in enumerate(tokens):
        if not is_token:
            out.append(text)
            continue
        key = text.lower()
        if key.startswith("m") and len(key) <= 2 and _is_minute(tokens, index):
            out.append(f"{minute:0{len(key)}d}")
        elif key.startswith("m") and len(key) > 2 and key not in ("mmm", "mmmm"):
            raise UncalculableError("month format")
        else:
            out.append(_date_part(key, day, hour, minute, second, twelve_hour, text))
    return "".join(out)


def _split_date_tokens(text: str) -> list[tuple[bool, str]]:
    tokens, position = [], 0
    for match in _DATE_TOKEN.finditer(text):
        if match.start() > position:
            tokens.extend((False, c) for c in _check_literals(text[position : match.start()]))
        tokens.append((True, match.group()))
        position = match.end()
    tokens.extend((False, c) for c in _check_literals(text[position:]))
    return tokens


def _check_literals(text: str) -> str:
    if any(c not in _LITERALS and c not in ".," for c in text):
        raise UncalculableError("date format")
    return text


def _is_minute(tokens: list[tuple[bool, str]], index: int) -> bool:
    """An "m" is minutes when it follows hours or precedes seconds."""
    before = [t.lower() for is_token, t in tokens[:index] if is_token]
    after = [t.lower() for is_token, t in tokens[index + 1 :] if is_token]
    return bool((before and before[-1] in ("h", "hh")) or (after and after[0] in ("s", "ss")))


def _date_part(
    key: str,
    day: dt.date,
    hour: int,
    minute: int,
    second: int,
    twelve_hour: bool,
    original: str,
) -> str:
    match key:
        case "yyyy":
            return f"{day.year:04d}"
        case "yy" | "y":
            return f"{day.year % 100:02d}"
        case "mmmm":
            return _MONTHS[day.month - 1]
        case "mmm":
            return _MONTHS[day.month - 1][:3]
        case "mm" | "m":
            return f"{day.month:0{len(key)}d}"
        case "dddd":
            return _DAYS[day.weekday()]
        case "ddd":
            return _DAYS[day.weekday()][:3]
        case "dd" | "d":
            return f"{day.day:0{len(key)}d}"
        case "hh" | "h":
            shown = (hour % 12 or 12) if twelve_hour else hour
            return f"{shown:0{len(key)}d}"
        case "ss" | "s":
            return f"{second:0{len(key)}d}"
        case "am/pm":
            return (
                ("AM" if hour < 12 else "PM")
                if original[0] == "A"
                else ("am" if hour < 12 else "pm")
            )
        case "a/p":
            return "A" if hour < 12 else "P"
    raise UncalculableError("date format")
