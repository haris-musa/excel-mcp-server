"""The kinds of conditional format rule and what each stores."""

from typing import Literal

RuleType = Literal[
    "color_scale",
    "data_bar",
    "icon_set",
    "cell_value",
    "formula",
    "top",
    "bottom",
    "above_average",
    "below_average",
    "duplicate",
    "unique",
    "contains_text",
    "not_contains_text",
    "begins_with",
    "ends_with",
    "date",
    "blanks",
    "no_blanks",
    "errors",
    "no_errors",
]
Period = Literal[
    "yesterday",
    "today",
    "tomorrow",
    "last7Days",
    "lastWeek",
    "thisWeek",
    "nextWeek",
    "lastMonth",
    "thisMonth",
    "nextMonth",
]
TextRuleType = Literal["containsText", "notContainsText", "beginsWith", "endsWith"]
TextOperator = Literal["containsText", "notContains", "beginsWith", "endsWith"]
CellRuleType = Literal["containsBlanks", "notContainsBlanks", "containsErrors", "notContainsErrors"]
ThresholdType = Literal["percent", "num", "percentile"]
IconSetName = Literal[
    "3Arrows",
    "3ArrowsGray",
    "3Flags",
    "3TrafficLights1",
    "3TrafficLights2",
    "3Signs",
    "3Symbols",
    "3Symbols2",
    "4Arrows",
    "4ArrowsGray",
    "4RedToBlack",
    "4Rating",
    "4TrafficLights",
    "5Arrows",
    "5ArrowsGray",
    "5Rating",
    "5Quarters",
]

# What Excel writes for each period; {c} is the top-left cell of the range.
PERIOD_FORMULAS: dict[Period, str] = {
    "yesterday": "FLOOR({c},1)=TODAY()-1",
    "today": "FLOOR({c},1)=TODAY()",
    "tomorrow": "FLOOR({c},1)=TODAY()+1",
    "last7Days": "AND(TODAY()-FLOOR({c},1)<=6,FLOOR({c},1)<=TODAY())",
    "lastWeek": "AND(TODAY()-ROUNDDOWN({c},0)>=(WEEKDAY(TODAY())),"
    "TODAY()-ROUNDDOWN({c},0)<(WEEKDAY(TODAY())+7))",
    "thisWeek": "AND(TODAY()-ROUNDDOWN({c},0)<=WEEKDAY(TODAY())-1,"
    "ROUNDDOWN({c},0)-TODAY()<=7-WEEKDAY(TODAY()))",
    "nextWeek": "AND(ROUNDDOWN({c},0)-TODAY()>(7-WEEKDAY(TODAY())),"
    "ROUNDDOWN({c},0)-TODAY()<(15-WEEKDAY(TODAY())))",
    "lastMonth": "AND(MONTH({c})=MONTH(EDATE(TODAY(),0-1)),YEAR({c})=YEAR(EDATE(TODAY(),0-1)))",
    "thisMonth": "AND(MONTH({c})=MONTH(TODAY()),YEAR({c})=YEAR(TODAY()))",
    "nextMonth": "AND(MONTH({c})=MONTH(EDATE(TODAY(),0+1)),YEAR({c})=YEAR(EDATE(TODAY(),0+1)))",
}
TEXT_RULES: dict[str, tuple[TextRuleType, TextOperator, str]] = {
    "contains_text": ("containsText", "containsText", "NOT(ISERROR(SEARCH({t},{c})))"),
    "not_contains_text": ("notContainsText", "notContains", "ISERROR(SEARCH({t},{c}))"),
    "begins_with": ("beginsWith", "beginsWith", "LEFT({c},LEN({t}))={t}"),
    "ends_with": ("endsWith", "endsWith", "RIGHT({c},LEN({t}))={t}"),
}
CELL_FORMULAS: dict[str, tuple[CellRuleType, str]] = {
    "blanks": ("containsBlanks", "LEN(TRIM({c}))=0"),
    "no_blanks": ("notContainsBlanks", "LEN(TRIM({c}))>0"),
    "errors": ("containsErrors", "ISERROR({c})"),
    "no_errors": ("notContainsErrors", "NOT(ISERROR({c}))"),
}
# Fields that every rule accepts apart from these are rejected when they do not fit the type.
FORMATTED = {"fill_color", "font_color"}
FIELDS = {
    "color_scale": {"colors"},
    "data_bar": {"colors"},
    "icon_set": {"icon_set", "thresholds", "threshold_type", "reverse", "icon_only"},
    "cell_value": {"operator", "values"} | FORMATTED,
    "formula": {"formula"} | FORMATTED,
    "top": {"count", "percent"} | FORMATTED,
    "bottom": {"count", "percent"} | FORMATTED,
    "above_average": {"std_dev", "include_equal"} | FORMATTED,
    "below_average": {"std_dev", "include_equal"} | FORMATTED,
    "duplicate": FORMATTED,
    "unique": FORMATTED,
    "contains_text": {"text"} | FORMATTED,
    "not_contains_text": {"text"} | FORMATTED,
    "begins_with": {"text"} | FORMATTED,
    "ends_with": {"text"} | FORMATTED,
    "date": {"period"} | FORMATTED,
    "blanks": FORMATTED,
    "no_blanks": FORMATTED,
    "errors": FORMATTED,
    "no_errors": FORMATTED,
}
ALWAYS = {"type", "stop_if_true", "priority"}
