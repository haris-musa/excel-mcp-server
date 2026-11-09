"""Excel number formats (US English): sections, conditions, digits, scientific, fractions, dates.

Formats for another language raise `UncalculableError` rather than guess.
"""

import re
from dataclasses import dataclass, field
from decimal import ROUND_HALF_UP, Decimal
from fractions import Fraction

from excel_mcp.calc.dateparts import LAST_SERIAL, day_of_week, month_day_year
from excel_mcp.calc.values import VALUE, FormulaError, UncalculableError

Item = tuple[str, str]  # (kind, text): lit, num, den, exp, general, date, elapsed

_COLORS = {"black", "blue", "cyan", "green", "magenta", "red", "white", "yellow"}
_CONDITION = re.compile(r"(<=|>=|<>|<|>|=)\s*(-?\d+\.?\d*(?:[eE][+-]?\d+)?)")
_ELAPSED = re.compile(r"h+|m+|s+", re.I)
_DATE_WORD = re.compile(
    r"(?i:yyyy|yyy|yy|y|mmmmm|mmmm|mmm|mm|m|dddd|ddd|dd|d|hh|h|ss|s|AM/PM|A/P)|ee|e"
)
_HAS_DATE = re.compile(r"(?i:[ydhsm]|AM/PM|A/P)|e|§")
_LITERALS = set("$-+/():!^&'~{}<>= ")
_SYSTEM_FORMATS = {"F800": "d mmmm, yyyy", "F400": "h:mm:ss AM/PM"}
_MONTHS = [
    "January", "February", "March", "April", "May", "June", "July", "August",
    "September", "October", "November", "December",
]  # fmt: skip
_DAYS = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]
_COMPARE = {
    "<": lambda a, b: a < b, ">": lambda a, b: a > b, "<=": lambda a, b: a <= b,
    ">=": lambda a, b: a >= b, "=": lambda a, b: a == b, "<>": lambda a, b: a != b,
}  # fmt: skip


@dataclass
class Section:
    condition: tuple[str, float] | None = None
    prefix: str = ""  # a currency symbol from a code such as [$€-407]
    items: list[Item] = field(default_factory=list)


def _split_sections(fmt: str) -> list[str]:
    parts, current, quoted, escaped = [], "", False, False
    for character in fmt:
        if not escaped and character == ";" and not quoted:
            parts.append(current)
            current = ""
            continue
        current += character
        if escaped:
            escaped = False
        elif character == "\\":
            escaped = True
        elif character == '"':
            quoted = not quoted
    parts.append(current)
    if len(parts) > 4:
        raise FormulaError(VALUE)
    return parts


def _section(text: str) -> Section:
    section = Section()
    system = ""

    def bracket(body: str) -> str:
        nonlocal system
        if body.lower() in _COLORS or re.fullmatch(r"color\s*\d+", body, re.I):
            return ""
        if condition := _CONDITION.fullmatch(body):
            section.condition = (condition[1], float(condition[2]))
            return ""
        if _ELAPSED.fullmatch(body):
            return f"\x00{body.lower()}\x00"
        if body.startswith("$"):
            symbol, _, code = body[1:].partition("-")
            if code.upper() in _SYSTEM_FORMATS:
                system = _SYSTEM_FORMATS[code.upper()]
            elif code and int(code, 16) & 0xFF != 9:
                raise UncalculableError("number format for another language")
            section.prefix += symbol
            return ""
        raise FormulaError(VALUE)

    stripped = re.sub(r'("[^"]*"|\\.)|\[([^\]]*)\]', lambda m: m[1] or bracket(m[2]), text)
    section.items = _items(system or stripped)
    return section


def _items(text: str) -> list[Item]:
    items: list[Item] = []
    plain = re.sub(r'"[^"]*"|\\.|general|[eE][+-]', "", text, flags=re.I)
    dated = bool(_HAS_DATE.search(plain))
    index = 0
    while index < len(text):
        character = text[index]
        rest = text[index:]
        if character == '"':
            end = text.find('"', index + 1)
            if end < 0:
                raise FormulaError(VALUE)
            items.append(("lit", text[index + 1 : end]))
            index = end + 1
        elif character == "\\":
            items.append(("lit", text[index + 1 : index + 2]))
            index += 2
        elif character == "_":
            items.append(("lit", " "))
            index += 2
        elif character == "*":
            index += 2  # a fill character repeats only to the width of a cell
        elif character == "\x00":
            end = text.index("\x00", index + 1)
            items.append(("elapsed", text[index + 1 : end]))
            index = end + 1
        elif rest[:7].lower() == "general":
            items.append(("general", ""))
            index += 7
        elif dated and (match := _DATE_WORD.match(rest)):
            items.append(("date", match[0]))
            index += len(match[0])
        elif character == "E" and rest[1:2] in ("+", "-"):
            items.append(("exp", rest[:2]))
            index += 2
        elif character == "/" and items and (match := re.match(r"/(\d+)", rest)):
            items.append(("num", "/"))
            items.append(("den", match[1]))
            index += len(match[0])
        elif character in "0#?.,%/@":
            kind = "lit" if dated and character in ",/%" else "num"
            items.append((kind, character))
            index += 1
        elif character in _LITERALS or character.isspace() or _plain_letter(character):
            items.append(("lit", character))
            index += 1
        elif character in "bBgG":
            raise UncalculableError("era in a number format")
        else:
            raise FormulaError(VALUE)
    return items


def _plain_letter(character: str) -> bool:
    """Letters that Excel shows as they are: all but the date and exponent codes and N."""
    return character.isalpha() and character.lower() not in "bgne"


def format_value(value: float | str, fmt: str) -> str:
    """The text of TEXT(value, fmt)."""
    sections = [_section(part) for part in _split_sections(fmt)]
    if isinstance(value, str):
        return _format_text(value, sections)
    section, minus = _choose(value, sections)
    kinds = {kind for kind, _ in section.items}
    if not section.items:
        return ""
    negative = value < 0 and minus
    if kinds & {"date", "elapsed"}:
        lone = _lone_minutes_or_seconds(section.items)
        if value < 0 and not lone:
            raise FormulaError(VALUE)
        return section.prefix + _date(value, section.items, lone, negative=value < 0)
    magnitude = abs(value)
    if ("num", "@") in section.items:
        if len(sections) > 1:
            raise FormulaError(VALUE)
        body = ("-" if value < 0 else "") + _general(magnitude)
    elif "exp" in kinds:
        if sum(kind == "exp" for kind, _ in section.items) > 1:
            raise FormulaError(VALUE)
        body = _scientific(magnitude, negative, section.items)
    elif ("num", "/") in section.items:
        body = _fraction(magnitude, negative, section.items)
    elif "general" in kinds:
        if any(kind == "num" and item in "0#?" for kind, item in section.items):
            raise FormulaError(VALUE)
        text = _general(magnitude)
        body = ("-" if negative else "") + "".join(
            text if kind == "general" else item for kind, item in section.items if kind != "num"
        )
    else:
        body = _digits(magnitude, negative, section.items)
    return section.prefix + body


def _general(magnitude: float) -> str:
    """Excel's General format as TEXT gives it: what fits in 11 characters."""
    if magnitude == 0:
        return "0"
    number = Decimal(f"{magnitude:.15g}")
    if number >= 1:
        digits = number.adjusted() + 1
        if digits <= 11:
            places = max(10 - digits, 0)
            rounded = number.quantize(Decimal(1).scaleb(-places), rounding=ROUND_HALF_UP)
            if rounded.adjusted() + 1 <= 11:
                text = f"{rounded:f}"
                return text.rstrip("0").rstrip(".") if "." in text else text
    else:
        plain = f"{number:f}"
        if len(plain) <= 11:
            return plain
        if number >= Decimal("0.0001"):
            rounded = number.quantize(Decimal(1).scaleb(-9), rounding=ROUND_HALF_UP)
            return f"{rounded:f}".rstrip("0").rstrip(".")
    exponent = number.adjusted()
    significant = 5 if abs(exponent) >= 99 else 6
    mantissa = number.scaleb(-exponent).quantize(Decimal(1).scaleb(1 - significant), ROUND_HALF_UP)
    if mantissa >= 10:
        exponent += 1
        mantissa = number.scaleb(-exponent).quantize(
            Decimal(1).scaleb(1 - significant), ROUND_HALF_UP
        )
    text = f"{mantissa:f}".rstrip("0").rstrip(".") if "." in f"{mantissa:f}" else f"{mantissa:f}"
    return f"{text}E{'-' if exponent < 0 else '+'}{abs(exponent):02d}"


def _choose(value: float, sections: list[Section]) -> tuple[Section, bool]:
    """The section for a number, and whether it shows the minus sign itself."""
    if len(sections) == 1:
        only = sections[0]
        if only.condition and not _COMPARE[only.condition[0]](value, only.condition[1]):
            raise FormulaError(VALUE)
        return only, True
    if any(section.condition for section in sections):
        for index, section in enumerate(sections[:3]):
            last = index == len(sections) - 1 or index == 2
            if section.condition is None and last:
                return section, True
            if section.condition and _COMPARE[section.condition[0]](value, section.condition[1]):
                return section, section.condition not in (("<", 0), ("<=", 0))
        raise FormulaError(VALUE)
    if value < 0:
        return sections[1], False
    if value == 0 and len(sections) >= 3:
        return sections[2], True
    return sections[0], True


def _format_text(text: str, sections: list[Section]) -> str:
    placeholder = ("num", "@")
    chosen = (
        sections[3]
        if len(sections) == 4
        else next((s for s in sections if placeholder in s.items), None)
    )
    if chosen is None:
        return text
    return chosen.prefix + "".join(
        text if item == placeholder else item[1]
        for item in chosen.items
        if item[0] in ("lit", "num")
        if item == placeholder or item[0] == "lit"
    )


def _slots(items: list[Item], dot: int | None) -> tuple[list[int], list[int]]:
    places = [i for i, (kind, item) in enumerate(items) if kind == "num" and item in "0#?"]
    return (
        [i for i in places if dot is None or i < dot],
        [i for i in places if dot is not None and i > dot],
    )


def _fill_integer(items: list[Item], slots: list[int], digits: str) -> tuple[dict[int, str], str]:
    chars: dict[int, str] = {}
    remaining = list(digits)
    for slot in reversed(slots):
        kind = items[slot][1]
        chars[slot] = remaining.pop() if remaining else {"0": "0", "?": " ", "#": ""}[kind]
    return chars, "".join(remaining)


def _fill_fraction(items: list[Item], slots: list[int], fraction: str) -> dict[int, str]:
    chars: dict[int, str] = {}
    for number, slot in enumerate(slots):
        kind = items[slot][1]
        trailing_zeros = not any(c != "0" for c in fraction[number:])
        if trailing_zeros and kind != "0":
            chars[slot] = "" if kind == "#" else " "
        else:
            chars[slot] = fraction[number] if number < len(fraction) else "0"
    return chars


def _group(chars: dict[int, str], slots: list[int], overflow: str) -> tuple[dict[int, str], str]:
    """Thousands separators after the digit that starts each group of three from the right."""
    total = len(overflow) + sum(chars[slot].isdigit() for slot in slots)
    seen = 0
    grouped: dict[int, str] = {}
    for slot in reversed(slots):
        char = chars[slot]
        if char.isdigit():
            seen += 1
            char = ("," if seen % 3 == 0 and seen < total else "") + char
        grouped[slot] = char
    extra = ""
    for char in reversed(overflow):
        seen += 1
        extra = ("," if seen % 3 == 0 and seen < total else "") + char + extra
    return grouped, extra


def _scale(items: list[Item]) -> tuple[int, int, bool]:
    """The power of 1000 to divide by, the power of 100 to multiply by, and whether to group."""
    percent = sum(item == ("num", "%") for item in items)
    places = [i for i, (kind, item) in enumerate(items) if kind == "num" and item in "0#?"]
    if not places:
        return 0, percent, False
    scale, cursor = 0, places[-1] + 1
    while cursor < len(items) and items[cursor] in (("num", ","), ("lit", " ")):
        scale += items[cursor] == ("num", ",")
        cursor += 1
    grouping = any(items[i] == ("num", ",") for i in range(places[0], places[-1]))
    return scale, percent, grouping


def _digits(magnitude: float, negative: bool, items: list[Item]) -> str:
    scale, percent, grouping = _scale(items)
    dot = next((i for i, item in enumerate(items) if item == ("num", ".")), None)
    int_slots, frac_slots = _slots(items, dot)
    shown = Decimal(f"{magnitude:.15g}") * Decimal(100) ** percent / Decimal(1000) ** scale
    rounded = shown.quantize(Decimal(1).scaleb(-len(frac_slots)), rounding=ROUND_HALF_UP)
    whole, _, fraction = f"{rounded:f}".partition(".")
    chars, overflow = _fill_integer(items, int_slots, whole.lstrip("0"))
    chars |= _fill_fraction(items, frac_slots, fraction)
    if grouping:
        grouped, overflow = _group(chars, int_slots, overflow)
        chars |= grouped
    out = ["-" if negative and rounded != 0 else ""]
    for index, (kind, item) in enumerate(items):
        if index in chars:
            out.append((overflow if int_slots and index == int_slots[0] else "") + chars[index])
        elif kind == "num" and item == "." and not int_slots:
            out.append(overflow + ".")
        elif kind == "lit" or (item in (".", "%") and kind == "num"):
            out.append(item)
        elif kind == "num" and item in "@/":
            raise UncalculableError("number format")
    return "".join(out)


def _scientific(magnitude: float, negative: bool, items: list[Item]) -> str:
    at = next(i for i, (kind, _) in enumerate(items) if kind == "exp")
    head, tail = items[:at], items[at:]
    dot = next((i for i, item in enumerate(head) if item == ("num", ".")), None)
    int_slots, frac_slots = _slots(head, dot)
    if not int_slots and not frac_slots:
        raise FormulaError(VALUE)
    width = max(len(int_slots), 1)
    number = Decimal(f"{magnitude:.15g}")
    exponent = 0 if number == 0 else (number.adjusted() // width) * width
    quantum = Decimal(1).scaleb(-len(frac_slots))
    mantissa = number.scaleb(-exponent).quantize(quantum, rounding=ROUND_HALF_UP)
    if mantissa >= Decimal(10) ** width:
        exponent += width
        mantissa = number.scaleb(-exponent).quantize(quantum, rounding=ROUND_HALF_UP)
    whole, _, fraction = f"{mantissa:f}".partition(".")
    chars, overflow = _fill_integer(head, int_slots, whole.lstrip("0"))
    if number == 0:
        chars = dict.fromkeys(int_slots if width > 1 else int_slots[-1:], "0")
    chars |= _fill_fraction(head, frac_slots, fraction)
    out = ["-" if negative and mantissa != 0 else ""]
    for index, (kind, item) in enumerate(head):
        if index in chars:
            out.append((overflow if int_slots and index == int_slots[0] else "") + chars[index])
        elif kind == "lit" or (item == "." and kind == "num"):
            out.append(item)
    sign = "-" if exponent < 0 else ("+" if tail[0][1][1] == "+" else "")
    digits = sum(item == ("num", "0") for item in tail[1:])
    if not digits:
        raise FormulaError(VALUE)
    out.append(tail[0][1][0] + sign + f"{abs(exponent):0{digits}d}")
    out.extend(item for kind, item in tail[1:] if kind == "lit")
    return "".join(out)


def _half_up(number: Fraction) -> int:
    return (number + Fraction(1, 2)).__floor__()


def _convergent(number: Fraction, limit: int) -> tuple[int, int]:
    """The last continued-fraction convergent of ``number`` with a denominator up to ``limit``."""
    previous, current = (0, 1), (1, 0)
    rest = number
    while True:
        term = rest.__floor__()
        candidate = (term * current[0] + previous[0], term * current[1] + previous[1])
        if candidate[1] > limit:
            return current
        previous, current = current, candidate
        if rest == term:
            return current
        rest = 1 / (rest - term)


def _fraction(magnitude: float, negative: bool, items: list[Item]) -> str:
    slash = items.index(("num", "/"))
    before, after = items[:slash], items[slash + 1 :]
    places = [i for i, (kind, item) in enumerate(before) if kind == "num" and item in "0#?"]
    gap = next((i for i in range(len(before) - 1, -1, -1) if before[i] == ("lit", " ")), None)
    whole_slots = [i for i in places if gap is not None and i < gap]
    top_slots = [i for i in places if gap is None or i > gap]
    fixed = next((item for kind, item in after if kind == "den"), None)
    bottom_width = len(fixed) if fixed else sum(1 for k, i in after if k == "num" and i in "0#?")
    number = Fraction(Decimal(f"{magnitude:.15g}"))
    whole = int(number) if whole_slots else 0
    rest = number - whole
    if fixed is not None:
        denominator = int(fixed)
        numerator = _half_up(rest * denominator)
    else:
        numerator, denominator = _convergent(rest, 10**bottom_width - 1)
    if whole_slots and numerator == denominator:
        whole, numerator = whole + 1, 0
    if numerator >= 2**31:
        raise FormulaError(VALUE)
    shown = numerator != 0 or not whole_slots
    top = (str(numerator) if shown else "").rjust(len(top_slots))
    out = ["-" if negative and (whole_slots or numerator) else ""]
    for index, (kind, item) in enumerate(before):
        if index == (whole_slots[0] if whole_slots else -1):
            out.append(str(whole) if whole or not shown else "")
        elif index == top_slots[0]:
            out.append(top)
        elif kind == "lit":
            out.append(item)
    out.append("/" if shown else " ")
    out.append((str(denominator) if shown else "").ljust(bottom_width))
    return "".join(out)


def _digits_of(count: int) -> str:
    """A whole number as Excel writes it: only its first 15 digits are kept."""
    text = str(count)
    if len(text) <= 15:
        return text
    kept = int(text[:15]) + (text[15] >= "5")
    return str(kept) + "0" * (len(text) - 15)


def _lone_minutes_or_seconds(items: list[Item]) -> bool:
    """[m], [mm], [s] or [ss] alone count units without Excel's limits on dates."""
    marks = [item for item in items if item[0] in ("date", "elapsed")]
    return len(marks) == 1 and marks[0][0] == "elapsed" and marks[0][1][0] in "ms"


def _date(value: float, items: list[Item], lone: bool = False, *, negative: bool = False) -> str:
    decimals = _subsecond_digits(items)
    units = 10**decimals
    total = _half_up(Fraction(Decimal(f"{abs(value):.15g}")) * 86400 * units)
    days, seconds = divmod(total, 86400 * units)
    if days > LAST_SERIAL and not lone:
        raise FormulaError(VALUE)
    if lone:
        days %= LAST_SERIAL + 1
    hour, rest = divmod(seconds, 3600 * units)
    minute, rest = divmod(rest, 60 * units)
    second, subsecond = divmod(rest, units)
    year, month, day = month_day_year(days)
    weekday = day_of_week(days)
    twelve = any(kind == "date" and item.upper() in ("AM/PM", "A/P") for kind, item in items)
    elapsed = {"h": total // (3600 * units), "m": total // (60 * units), "s": total // units}
    fraction_digits = iter(f"{subsecond:0{decimals}d}")
    out: list[str] = []
    for index, (kind, item) in enumerate(items):
        key = item.lower()
        if kind == "lit" or (kind, item) == ("num", "."):
            out.append(item)
        elif kind == "elapsed":
            out.append(f"{_digits_of(elapsed[key[0]]):0>{len(key)}}")
        elif kind == "num" and item == "0" and decimals:
            out.append(next(fraction_digits, "0"))
        elif kind == "date":
            is_minute = key in ("m", "mm") and _is_minute(items, index)
            out.append(
                _date_part(
                    key,
                    item,
                    is_minute,
                    (year, month, day, weekday),
                    (hour, minute, second),
                    twelve,
                )
            )
        else:
            raise FormulaError(VALUE)
    return ("-" if negative else "") + "".join(out)


def _subsecond_digits(items: list[Item]) -> int:
    for index, (kind, item) in enumerate(items):
        if kind in ("date", "elapsed") and item.lower() in ("s", "ss"):
            rest = items[index + 1 :]
            if rest[:1] == [("num", ".")]:
                return sum(1 for entry in rest[1:] if entry == ("num", "0"))
    return 0


def _is_minute(items: list[Item], index: int) -> bool:
    def units(group: list[Item]) -> list[str]:
        return [i[1].lower()[0] for i in group if i[0] in ("date", "elapsed")]

    before, after = units(items[:index]), units(items[index + 1 :])
    return bool((before and before[-1] == "h") or (after and after[0] == "s"))


def _date_part(
    key: str,
    original: str,
    minute_token: bool,
    calendar: tuple[int, int, int, int],
    clock: tuple[int, int, int],
    twelve: bool,
) -> str:
    year, month, day, weekday = calendar
    hour, minute, second = clock
    if minute_token:
        return f"{minute:0{len(key)}d}"
    match key:
        case "yyyy" | "yyy" | "e" | "ee":
            return f"{year:04d}"
        case "yy" | "y":
            return f"{year % 100:02d}"
        case "mmmmm":
            return _MONTHS[month - 1][0]
        case "mmmm":
            return _MONTHS[month - 1]
        case "mmm":
            return _MONTHS[month - 1][:3]
        case "mm" | "m":
            return f"{month:0{len(key)}d}"
        case "dddd":
            return _DAYS[weekday]
        case "ddd":
            return _DAYS[weekday][:3]
        case "dd" | "d":
            return f"{day:0{len(key)}d}"
        case "hh" | "h":
            return f"{(hour % 12 or 12) if twelve else hour:0{len(key)}d}"
        case "ss" | "s":
            return f"{second:0{len(key)}d}"
        case "am/pm":
            pair = ("AM", "PM") if original[0] == "A" else ("am", "pm")
            return pair[hour >= 12]
        case "a/p":
            pair = ("A", "P") if original[0] == "A" else ("a", "p")
            return pair[hour >= 12]
    raise UncalculableError("date format")
