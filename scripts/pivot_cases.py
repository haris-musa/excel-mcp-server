"""Source data and PivotTable cases for the Excel golden fixture (see pivot_golden.py).

Each case is the arguments of create_pivot_table beyond the source and target. The data has
blank labels and values, zeros, a repeated label in another case ("east"), and a unique
"Noise" column, so that sorting by value never ties (Excel orders ties its own way).
"""

import datetime as dt
import itertools
from typing import Any

SOURCE: list[list[Any]] = [
    ["Region", "Product", "Channel", "Date", "Units", "Price", "Noise", "Qty"],
    ["North", "Apples", "Web", dt.datetime(2024, 1, 15), 10, 1.5, 11.5, 3],
    ["South", "Apples", "Shop", dt.datetime(2024, 2, 20), 5, 1.5, 7.25, 8],
    ["North", "Pears", "Web", dt.datetime(2024, 5, 3), 7, 2.0, 3.125, 12],
    ["South", "Pears", "Web", dt.datetime(2025, 1, 9), 3, 2.0, 21.5, 5],
    ["East", "Apples", "Shop", dt.datetime(2025, 3, 30), 12, 1.5, 4.75, 9],
    ["East", "Pears", "Shop", dt.datetime(2025, 7, 1), 8, 2.0, 16.0, 14],
    ["North", "Apples", "Shop", dt.datetime(2025, 7, 14), 6, 1.5, 9.0625, 1],
    ["East", "Kiwis", "Web", dt.datetime(2025, 7, 20), None, 3.0, 2.5, 6],
    ["South", "Kiwis", "Web", dt.datetime(2025, 8, 2), 9, 3.0, 18.75, 10],
    ["North", "Kiwis", "Shop", dt.datetime(2025, 8, 5), 11, None, 5.5, 4],
    ["West", "Apples", "Web", dt.datetime(2026, 9, 5), 0, 1.5, 13.25, 2],
    [None, "Pears", "Shop", dt.datetime(2026, 9, 6), 1, 2.0, 1.875, 11],
    ["east", "Kiwis", "Shop", dt.datetime(2026, 1, 1), 20, 3.0, 24.5, 7],
    ["West", "Kiwis", "Web", dt.datetime(2024, 2, 29), 14, 3.5, 6.375, 13],
]
SOURCE_RANGE = "A1:H15"
UNITS = [{"field": "Units"}]
TWO = [{"field": "Units"}, {"field": "Price", "function": "average"}]
SHAPES: list[tuple[list[str], list[str]]] = [
    (["Region"], []),
    (["Region", "Product"], []),
    (["Region"], ["Product"]),
    (["Region", "Product"], ["Channel"]),
    (["Region"], ["Product", "Channel"]),
]
BASED = [
    "percent_of_parent",
    "running_total",
    "percent_running_total",
    "rank_ascending",
    "rank_descending",
    "difference_from",
    "percent_difference_from",
    "percent_of",
]
PLAIN = [
    "percent_of_total",
    "percent_of_row",
    "percent_of_column",
    "percent_of_parent_row",
    "percent_of_parent_column",
]


def cases() -> list[dict[str, Any]]:
    found = [*_layouts(), *_figures(), *_sorting(), *_items(), *_groups(), *_calculated()]
    for number, case in enumerate(found, 1):
        case["name"] = f"c{number}"
    return found


def _layouts() -> list[dict[str, Any]]:
    found = []
    for layout, subtotals, values_in in itertools.product(
        ("compact", "outline", "tabular"), (True, False), ("columns", "rows")
    ):
        for rows, columns in SHAPES:
            values = TWO if values_in == "rows" or len(rows) + len(columns) > 2 else UNITS
            found.append(
                {
                    "rows": rows,
                    "columns": columns,
                    "values": values,
                    "layout": layout,
                    "subtotals": subtotals,
                    "values_in": values_in,
                }
            )
    return found


def _figures() -> list[dict[str, Any]]:
    found = []
    for rows, columns in SHAPES[:4]:
        for show in PLAIN:
            found.append(_shown(rows, columns, {"show_as": show}))
        for field in [*rows, *columns]:
            for show in BASED:
                found.append(_shown(rows, columns, {"show_as": show, "base_field": field}))
    found.append(
        _shown(
            ["Region"],
            ["Product"],
            {"show_as": "difference_from", "base_field": "Product", "base_item": "Pears"},
        )
    )
    found.append(
        _shown(
            ["Region"],
            ["Product"],
            {"show_as": "percent_of", "base_field": "Product", "base_item": "Kiwis"},
        )
    )
    found.append(
        {
            "rows": ["Region", "Product"],
            "values": [
                {
                    "field": "Price",
                    "function": "average",
                    "show_as": "percent_running_total",
                    "base_field": "Product",
                }
            ],
        }
    )
    found.append(
        {
            "rows": ["Region"],
            "columns": ["Product"],
            "values": [
                {"field": "Units"},
                {"field": "Units", "show_as": "percent_of_total"},
                {"field": "Price", "function": "max", "number_format": "$#,##0.00"},
            ],
        }
    )
    return found


def _shown(rows: list[str], columns: list[str], value: dict[str, Any]) -> dict[str, Any]:
    return {"rows": rows, "columns": columns, "values": [{"field": "Units", **value}]}


def _sorting() -> list[dict[str, Any]]:
    found = []
    for direction in ("ascending", "descending"):
        for rows, columns in SHAPES:
            fields = [{"field": field, "sort": direction} for field in [*rows[:1], *columns[:1]]]
            found.append({"rows": rows, "columns": columns, "values": UNITS, "fields": fields})
            by_value = [{"field": rows[0], "sort": direction, "sort_by": "Sum of Noise"}]
            if columns:
                by_value.append({"field": columns[0], "sort": direction, "sort_by": "Sum of Noise"})
            found.append(
                {
                    "rows": rows,
                    "columns": columns,
                    "values": [{"field": "Noise"}, *UNITS],
                    "fields": by_value,
                }
            )
    return found


def _items() -> list[dict[str, Any]]:
    return [
        {
            "rows": ["Region"],
            "values": UNITS,
            "fields": [{"field": "Region", "show_items": ["north", "South", "(blank)"]}],
        },
        {
            "rows": ["Region", "Product"],
            "values": UNITS,
            "fields": [{"field": "Product", "show_items": ["Apples", "Kiwis"]}],
        },
        {
            "rows": ["Region"],
            "columns": ["Product"],
            "values": UNITS,
            "fields": [{"field": "Product", "show_items": ["Pears"]}],
        },
        {
            "rows": ["Region"],
            "filters": ["Product"],
            "values": UNITS,
            "fields": [{"field": "Product", "show_items": ["Pears"]}],
        },
        {
            "rows": ["Region"],
            "filters": ["Product"],
            "values": UNITS,
            "fields": [{"field": "Product", "show_items": ["Pears", "Kiwis"]}],
        },
        {
            "rows": ["Region"],
            "filters": ["Product", "Channel"],
            "values": UNITS,
            "fields": [
                {"field": "Product", "show_items": ["Pears"]},
                {"field": "Channel", "show_items": ["Web"]},
            ],
        },
        {
            "rows": ["Region"],
            "filters": ["Date"],
            "values": UNITS,
            "fields": [{"field": "Date", "show_items": ["2024-01-15"]}],
        },
        {"rows": ["Region"], "filters": ["Channel"], "values": UNITS},
    ]


def _groups() -> list[dict[str, Any]]:
    found = []
    for units in (
        ["years"],
        ["quarters"],
        ["months"],
        ["days"],
        ["years", "months"],
        ["years", "quarters"],
        ["years", "quarters", "months"],
    ):
        group = {"field": "Date", "group_dates": units}
        found.append({"rows": ["Date"], "values": UNITS, "fields": [group]})
        found.append({"rows": ["Region"], "columns": ["Date"], "values": UNITS, "fields": [group]})
    found.append(
        {
            "rows": ["Date"],
            "values": UNITS,
            "fields": [{"field": "Date", "group_dates": ["years", "months"], "sort": "descending"}],
        }
    )
    found.append(
        {
            "rows": ["Date", "Region"],
            "values": UNITS,
            "subtotals": False,
            "layout": "outline",
            "fields": [{"field": "Date", "group_dates": ["years", "quarters"]}],
        }
    )
    found.append(
        {
            "rows": ["Date"],
            "values": [{"field": "Units", "show_as": "running_total", "base_field": "Months"}],
            "fields": [{"field": "Date", "group_dates": ["years", "months"]}],
        }
    )
    for group in (
        {"by": 5, "start": 0, "end": 20},
        {"by": 4},
        {"by": 4, "start": 3, "end": 12},
        {"by": 10, "end": 30},
        {"by": 2.5, "start": 0, "end": 22},
    ):
        found.append(
            {
                "rows": ["Qty"],
                "values": [{"field": "Price"}],
                "fields": [{"field": "Qty", "group_numbers": group}],
            }
        )
    found.append(
        {
            "rows": ["Qty"],
            "values": [{"field": "Price"}],
            "fields": [
                {
                    "field": "Qty",
                    "group_numbers": {"by": 5, "start": 0, "end": 20},
                    "sort": "descending",
                }
            ],
        }
    )
    found.append(
        {
            "rows": ["Region"],
            "columns": ["Qty"],
            "values": UNITS,
            "fields": [{"field": "Qty", "group_numbers": {"by": 5}}],
        }
    )
    return found


def _calculated() -> list[dict[str, Any]]:
    revenue = [{"name": "Revenue", "formula": "Units*Price"}]
    return [
        {"rows": ["Region"], "values": [{"field": "Revenue"}], "calculated_fields": revenue},
        {
            "rows": ["Region", "Product"],
            "columns": ["Channel"],
            "values": [{"field": "Revenue"}, {"field": "Units"}],
            "calculated_fields": revenue,
        },
        {
            "rows": ["Region"],
            "columns": ["Product"],
            "values": [{"field": "Revenue", "show_as": "percent_of_total"}],
            "calculated_fields": revenue,
        },
        {
            "rows": ["Product"],
            "values": [{"field": "Ratio"}],
            "calculated_fields": [{"name": "Ratio", "formula": "Units/Price"}],
        },
        {
            "rows": ["Region"],
            "values": [{"field": "B"}],
            "calculated_fields": [
                {"name": "A", "formula": "=Units*2"},
                {"name": "B", "formula": "A+Price"},
            ],
        },
        {
            "rows": ["Region", "Product"],
            "values": [{"field": "Flag"}],
            "calculated_fields": [{"name": "Flag", "formula": "IF(Units>8,1,0)"}],
        },
        {
            "rows": ["Region"],
            "values": [{"field": "Z"}],
            "calculated_fields": [{"name": "Z", "formula": "Units/(Price-Price)"}],
        },
        {
            "rows": ["Region"],
            "values": [{"field": "Q"}],
            "calculated_fields": [{"name": "Q", "formula": "'Units' * 2 + (Price^2)/10"}],
        },
    ]
