"""The searches behind MATCH, VLOOKUP, XMATCH and friends, step for step as Excel runs them.

Excel's approximate searches are bisections that never check the data is sorted, so on
unsorted data, or data of mixed types, the answer is whatever the bisection lands on;
these reproduce it.
"""

from excel_mcp.calc.criteria import has_wildcard
from excel_mcp.calc.values import ERROR_NUMBERS, ExcelError, Scalar, compare
from excel_mcp.calc.wildcard import Wildcard

# "Smaller" and "larger" refer to the value relative to the key, whatever the sort order.
SMALLER, LARGER = -1, 1


def lookup_key(key: Scalar) -> Scalar:
    """MATCH and the VLOOKUP family read a blank key as 0."""
    return 0.0 if key is None else key


def _rank(value: Scalar) -> int:
    """Excel's order of types: numbers, text, logicals, errors, and last of all blank cells."""
    if value is None:
        return 4
    if isinstance(value, bool):
        return 2
    if isinstance(value, str):
        return 1
    return 3 if isinstance(value, ExcelError) else 0


def _order(value: Scalar, key: Scalar) -> int:
    rank, key_rank = _rank(value), _rank(key)
    if rank == key_rank == 3:
        return compare(_error_number(value), _error_number(key))
    if rank != key_rank or rank == 4:
        return (rank > key_rank) - (rank < key_rank)
    return compare(value, key)


def _error_number(error: Scalar) -> float:
    return float(ERROR_NUMBERS.get(str(error), 0))


def _same_type(value: Scalar, key: Scalar) -> bool:
    return value is not None and _rank(value) == _rank(key)


def exact(
    key: Scalar,
    values: list[Scalar],
    *,
    wildcards: bool,
    from_end: bool = False,
    blank_matches: bool = False,
) -> int | None:
    """0-based index of the first (or last) value equal to the key."""
    order = range(len(values) - 1, -1, -1) if from_end else range(len(values))
    if key is None and blank_matches:
        return next((i for i in order if values[i] is None), None)
    if wildcards and isinstance(key, str) and has_wildcard(key):
        pattern = Wildcard(key)
        return next(
            (i for i in order if isinstance(v := values[i], str) and pattern.fullmatch(v)), None
        )
    return next(
        (i for i in order if _same_type(values[i], key) and _order(values[i], key) == 0), None
    )


def last_not_above(key: Scalar, values: list[Scalar]) -> int | None:
    """MATCH type 1 and approximate VLOOKUP: the last value up to the key, by bisection.

    A value of another type cannot be compared, so the bisection judges by the next value
    of the key's type to its right, and takes the left half when there is none.
    """
    low, high, best = 0, len(values) - 1, None
    while low <= high:
        mid = (low + high) // 2
        if not _same_type(values[mid], key):
            mid = next((i for i in range(mid + 1, high + 1) if _same_type(values[i], key)), -1)
            if mid < 0:
                high = (low + high) // 2 - 1
                continue
        order = _order(values[mid], key)
        if order == 0:
            while mid + 1 < len(values) and _same_type(values[mid + 1], key):
                if _order(values[mid + 1], key) != 0:
                    break
                mid += 1
            return mid
        if order < 0:
            best, low = mid, mid + 1
        else:
            high = mid - 1
    return best


def first_not_below(key: Scalar, values: list[Scalar]) -> int | None:
    """MATCH type -1: the last value down to the key, scanning until a value falls below it."""
    best = None
    for index, value in enumerate(values):
        if not _same_type(value, key):
            continue
        order = _order(value, key)
        if order == 0:
            return index
        if order < 0:
            break
        best = index
    return best


def nearest(key: Scalar, values: list[Scalar], mode: int, *, from_end: bool) -> int | None:
    """XMATCH by a linear search: the key, else the nearest smaller (-1) or larger (1) value."""
    order = range(len(values) - 1, -1, -1) if from_end else range(len(values))
    best = None
    for i in order:
        value = values[i]
        difference = _order(value, key)
        if difference == 0:
            return i
        closer = difference < 0 if mode == SMALLER else difference > 0
        if closer and (best is None or _order(value, values[best]) * mode < 0):
            best = i
    return best


def bisect(key: Scalar, values: list[Scalar], *, nearest: int, descending: bool) -> int | None:
    """XMATCH and XLOOKUP binary search: the key, else the nearest smaller or larger value.

    Duplicates of the key resolve to the first of them, or to the last in a descending search.
    """
    low, high, best, found = 0, len(values) - 1, None, None
    while low <= high:
        mid = (low + high) // 2
        order = _order(values[mid], key)
        if order == 0:
            found = mid
            if descending:
                low = mid + 1
            else:
                high = mid - 1
            continue
        key_is_after = order > 0 if descending else order < 0
        if key_is_after:
            if nearest == (LARGER if descending else SMALLER):
                best = mid
            low = mid + 1
        else:
            if nearest == (SMALLER if descending else LARGER):
                best = mid
            high = mid - 1
    return best if found is None else found
