"""Excel wildcard patterns (* ? ~) matched without backtracking."""

Segment = list[str | None]


class Wildcard:
    """A case-insensitive pattern with * ? and ~ escapes.

    Split at each "*" into segments of literal characters and "?" (None), each
    segment is placed at its leftmost position after the previous one, which is
    always enough. Segments are found with a bit-parallel scan, so matching reads
    the text once, however long the pattern is and however many stars it has.
    """

    def __init__(self, pattern: str) -> None:
        self._segments: list[Segment] = [[]]
        index = 0
        while index < len(pattern):
            character = pattern[index]
            if character == "~" and index + 1 < len(pattern):
                index += 1
                self._segments[-1].append(pattern[index].lower())
            elif character == "*":
                self._segments.append([])
            else:
                self._segments[-1].append(None if character == "?" else character.lower())
            index += 1

    def fullmatch(self, text: str) -> bool:
        chars = [character.lower() for character in text]
        first, *rest = self._segments
        if not rest:
            return len(chars) == len(first) and _fits(first, chars, 0)
        last = rest.pop()
        end = len(chars) - len(last)
        if end < len(first) or not _fits(first, chars, 0) or not _fits(last, chars, end):
            return False
        return _place(rest, chars, len(first), end) is not None

    def search(self, text: str, begin: int) -> int | None:
        """The start of the first match at or after ``begin``, or None."""
        chars = [character.lower() for character in text]
        first, *rest = self._segments
        start = _find(first, chars, begin, len(chars))
        # If any start works the earliest does, because it leaves the most text for the rest.
        if start is None or _place(rest, chars, start + len(first), len(chars)) is None:
            return None
        return start


def _place(segments: list[Segment], chars: list[str], position: int, end: int) -> int | None:
    for segment in segments:
        found = _find(segment, chars, position, end)
        if found is None:
            return None
        position = found + len(segment)
    return position


def _fits(segment: Segment, chars: list[str], at: int) -> bool:
    return all(want is None or want == chars[at + i] for i, want in enumerate(segment))


def _find(segment: Segment, chars: list[str], start: int, end: int) -> int | None:
    """The first position from ``start`` where the segment fits before ``end``.

    Bit i of ``state`` says whether the first i + 1 characters of the segment end at the
    current character (the Shift-And algorithm), so each character costs one step.
    """
    if not segment:
        return start if start <= end else None
    anything = sum(1 << i for i, want in enumerate(segment) if want is None)
    masks: dict[str, int] = {}
    for i, want in enumerate(segment):
        if want is not None:
            masks[want] = masks.get(want, anything) | 1 << i
    done = 1 << (len(segment) - 1)
    state = 0
    for at in range(start, end):
        state = ((state << 1) | 1) & masks.get(chars[at], anything)
        if state & done:
            return at - len(segment) + 1
    return None
