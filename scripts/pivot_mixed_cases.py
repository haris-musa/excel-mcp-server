"""Sources with mixed types, errors, times and formulas, and PivotTable cases over them.

Part of the Excel golden fixture (see pivot_golden.py and pivot_cases.py).
"""

import datetime as dt
import itertools
from typing import Any

MIXED: list[list[Any]] = [
    ["Mix", "Group", "Val", "Num", "Noise"],
    ["b", "x", 1, 10, 3.5],
    [3, "y", 2, 20, 7.25],
    [dt.datetime(2024, 1, 15), "z", 3.5, 30, 1.125],
    [True, "x", "t", 40, 9.5],
    ["#N/A", "y", True, 50, 2.75],
    [None, "z", None, None, 6.0],
    ["a", "x", 4, 5, 12.5],
    [1.5, "y", 5, 6, 4.25],
    ["B", "z", 6, 7, 8.125],
    [10, "x", "u", 8, 10.5],
    [dt.datetime(2024, 2, 20), "y", 7, 9, 5.75],
    [False, "z", 8, 11, 13.0],
    ["#DIV/0!", "x", 9, 12, 0.625],
    ["c", "y", 10, 13, 11.25],
    [dt.time(6, 0), "z", 11, 14, 14.5],
    [3, "x", 12, 15, 15.75],
    ["3", "z", 13, 16, 16.5],
]
"""Row labels of every type, and a value field of numbers, text, booleans and blanks."""
ERRORS: list[list[Any]] = [
    ["Key", "Sub", "Val", "Noise"],
    ["z", "p", "#NUM!", 0.5],
    ["a", "p", 1, 1.5],
    ["a", "q", "#N/A", 2.25],
    ["b", "p", 5, 3.125],
    ["b", "p", 6, 4.5],
    ["c", "q", "#DIV/0!", 5.75],
    ["c", "p", "#N/A", 6.0],
    ["d", "q", "t", 7.25],
    ["d", "p", True, 8.5],
    ["e", "q", None, 9.0],
    ["f", "p", 2, 10.5],
    ["f", "q", "#VALUE!", 11.25],
]
"""Errors in a value field; the first record (z) holds an error that comes last in the table."""
_DAY = dt.datetime
TIMES: list[list[Any]] = [
    ["Stamp", "Time", "Dur", "Day", "Kind", "Val"],
    [_DAY(2024, 1, 15, 9, 30), dt.time(6, 0), dt.timedelta(hours=36), _DAY(2024, 1, 15), "a", 1],
    [_DAY(2024, 1, 15, 18), dt.time(12, 30), dt.timedelta(hours=1.5), _DAY(2024, 2, 20), "b", 2],
    [_DAY(2024, 2, 20), dt.time(6, 0), dt.timedelta(days=2, hours=3), _DAY(2024, 1, 15), "a", 4],
    [_DAY(2024, 3, 1, 12), None, dt.timedelta(minutes=45), _DAY(2024, 3, 1), "b", 8],
    [None, dt.time(23, 15), dt.timedelta(hours=1.5), _DAY(2024, 2, 20), "a", 16],
]
"""Times, durations, and dates with and without times; some blanks."""
FORMULAS: list[list[Any]] = [
    ["Key", "Val", "Flag"],
    ['="a"&"1"', "=1+1", "=TRUE"],
    ['=IF(1>0,"a1","b")', "=2*3", "=1>2"],
    ["b", "=SUM(1,2,3)", True],
    ['="c"', "=1/0", "=NA()"],
    ["b", "=NA()", False],
    ["a1", 10, '="x"'],
    ["=DATE(2024,1,15)", "=1+2", None],
    ['=""', 4, 5],
]
"""Formula cells, which have no stored results until the calculator gives them."""
SOURCES: dict[str, list[list[Any]]] = {
    "Mixed": MIXED,
    "Errors": ERRORS,
    "Times": TIMES,
    "Formulas": FORMULAS,
}


def _all(field: str) -> list[dict[str, Any]]:
    return [
        {"field": field, "function": function}
        for function in ("sum", "count", "average", "min", "max")
    ]


def _count(field: str) -> dict[str, Any]:
    return {"field": field, "function": "count"}


def _on(source: str, found: list[dict[str, Any]]) -> list[dict[str, Any]]:
    return [{**case, "source": source} for case in found]


def mixed_cases() -> list[dict[str, Any]]:
    return [*_mixed(), *_errors(), *_times(), *_formulas()]


def _mixed() -> list[dict[str, Any]]:
    val = {"field": "Val"}
    calc = [{"name": "Calc", "formula": "Val*2+Num"}]
    found: list[dict[str, Any]] = [
        {"rows": ["Mix"], "values": _all("Val")},
        {"rows": ["Mix"], "values": _all("Num")},
        {"rows": ["Group"], "values": _all("Val")},
        {"rows": ["Group"], "columns": ["Mix"], "values": [val]},
        {"rows": ["Mix"], "columns": ["Group"], "values": [val, {"field": "Num"}]},
        {"rows": ["Group", "Mix"], "values": [val, _count("Val")]},
        {"rows": ["Mix", "Group"], "values": [val, _count("Val")], "subtotals": False},
        {"rows": ["Group"], "filters": ["Mix"], "values": [val]},
        {
            "rows": ["Group"],
            "filters": ["Mix"],
            "values": [val],
            "fields": [{"field": "Mix", "show_items": ["a", "TRUE"]}],
        },
        {
            "rows": ["Mix"],
            "values": [val],
            "fields": [{"field": "Mix", "show_items": ["a", "(blank)", "TRUE", "#N/A", "3"]}],
        },
        {
            "rows": ["Mix"],
            "values": [{"field": "Val", "number_format": "0.00"}],
            "fields": [{"field": "Mix", "sort": "descending"}],
        },
        {
            "rows": ["Mix"],
            "values": [{"field": "Noise"}, val],
            "fields": [{"field": "Mix", "sort": "ascending", "sort_by": "Sum of Noise"}],
        },
        {
            "rows": ["Mix"],
            "values": [{"field": "Noise"}, val],
            "fields": [{"field": "Mix", "sort": "descending", "sort_by": "Sum of Noise"}],
        },
        {
            "rows": ["Mix"],
            "values": [val, {"field": "Num", "function": "average"}],
            "values_in": "rows",
        },
        {"rows": ["Group"], "values": [{"field": "Calc"}], "calculated_fields": calc},
        {"rows": ["Mix"], "values": [{"field": "Calc"}], "calculated_fields": calc},
    ]
    for layout, subtotals in itertools.product(("compact", "outline", "tabular"), (True, False)):
        # Excel writes the captions of subtotals of dates in the host's date format.
        rows = ["Group", "Mix"] if subtotals else ["Mix", "Group"]
        found.append({"rows": rows, "values": [val], "layout": layout, "subtotals": subtotals})
    for show in ("percent_of_total", "percent_of_row", "percent_of_column"):
        found.append({"rows": ["Mix"], "columns": ["Group"], "values": [{**val, "show_as": show}]})
    for show in (
        "running_total",
        "rank_ascending",
        "difference_from",
        "percent_of",
        "percent_of_parent",
    ):
        found.append(
            {
                "rows": ["Group", "Mix"],
                "values": [{"field": "Num", "show_as": show, "base_field": "Mix"}],
            }
        )
    return _on("Mixed", found)


def _errors() -> list[dict[str, Any]]:
    val = {"field": "Val"}
    found: list[dict[str, Any]] = [
        {"rows": ["Key"], "values": _all("Val")},
        {"rows": ["Sub"], "values": _all("Val")},
        {"rows": ["Key"], "columns": ["Sub"], "values": [val, _count("Val")]},
        {
            "rows": ["Sub"],
            "columns": ["Key"],
            "values": [val, {"field": "Val", "function": "average"}],
        },
        {"rows": ["Key", "Sub"], "values": _all("Val")},
        {"rows": ["Sub", "Key"], "values": [val, _count("Val")]},
        {"rows": ["Key", "Sub"], "values": _all("Val"), "values_in": "rows"},
        {"rows": ["Key"], "values": [{"field": "Noise"}, _count("Val")]},
        {
            "rows": ["Key"],
            "filters": ["Sub"],
            "values": [val],
            "fields": [{"field": "Sub", "show_items": ["p"]}],
        },
        {
            "rows": ["Key"],
            "values": [{"field": "Noise"}, val],
            "fields": [{"field": "Key", "sort": "descending", "sort_by": "Sum of Noise"}],
        },
        {
            "rows": ["Key"],
            "values": [{"field": "Calc"}],
            "calculated_fields": [{"name": "Calc", "formula": "Val*2"}],
        },
        {
            "rows": ["Key", "Sub"],
            "values": [{"field": "Calc"}, {"field": "Noise"}],
            "calculated_fields": [{"name": "Calc", "formula": "Val+Noise"}],
        },
    ]
    for layout, subtotals in itertools.product(("compact", "tabular"), (True, False)):
        found.append(
            {"rows": ["Key", "Sub"], "values": [val], "layout": layout, "subtotals": subtotals}
        )
    for show in ("percent_of_total", "percent_of_column"):
        found.append({"rows": ["Key"], "columns": ["Sub"], "values": [{**val, "show_as": show}]})
    for show in (
        "running_total",
        "difference_from",
        "percent_of",
        "rank_ascending",
        "rank_descending",
        "percent_running_total",
        "percent_difference_from",
        "percent_of_parent",
    ):
        found.append(
            {"rows": ["Key"], "values": [{"field": "Val", "show_as": show, "base_field": "Key"}]}
        )
    found += [
        {"rows": ["Val"], "values": [{"field": "Noise"}]},
        {"rows": ["Val"], "columns": ["Sub"], "values": [{"field": "Noise"}]},
        {"rows": ["Key", "Sub"], "values": [val], "values_in": "rows", "subtotals": False},
        {
            "rows": ["Key"],
            "values": [{"field": "Val", "show_as": "percent_of_parent_row"}],
            "columns": ["Sub"],
        },
    ]
    return _on("Errors", found)


def _times() -> list[dict[str, Any]]:
    value = [{"field": "Val"}]
    found: list[dict[str, Any]] = [
        {"rows": ["Stamp"], "values": value},
        {"rows": ["Time"], "values": value},
        {"rows": ["Dur"], "values": value},
        {"rows": ["Day"], "values": value},
        {"rows": ["Kind"], "values": _all("Dur")},
        {"rows": ["Kind"], "values": _all("Time")},
        {"rows": ["Kind"], "values": _all("Stamp")},
        {"rows": ["Kind"], "values": _all("Day")},
        {"rows": ["Kind"], "columns": ["Time"], "values": value},
        {"rows": ["Kind"], "columns": ["Dur"], "values": value},
        {"rows": ["Day", "Kind"], "values": [{"field": "Dur"}, *value]},
        {"rows": ["Kind"], "filters": ["Stamp"], "values": value},
        {
            "rows": ["Stamp"],
            "values": value,
            "fields": [{"field": "Stamp", "sort": "descending"}],
        },
        {"rows": ["Time"], "values": [{"field": "Val", "show_as": "percent_of_total"}]},
        {
            "rows": ["Dur"],
            "values": [{"field": "Val", "show_as": "running_total", "base_field": "Dur"}],
        },
        {"rows": ["Kind"], "columns": ["Stamp"], "values": value, "subtotals": False},
        {
            "rows": ["Day"],
            "values": value,
            "fields": [{"field": "Day", "group_dates": ["years", "months"]}],
        },
    ]
    return _on("Times", found)


def _formulas() -> list[dict[str, Any]]:
    val = {"field": "Val"}
    found: list[dict[str, Any]] = [
        {"rows": ["Key"], "values": [val, _count("Val")]},
        {"rows": ["Key"], "values": _all("Val")},
        {"rows": ["Flag"], "values": [val, _count("Val")]},
        {"rows": ["Key", "Flag"], "values": [val]},
        {"rows": ["Key"], "columns": ["Flag"], "values": [val]},
    ]
    return _on("Formulas", found)
