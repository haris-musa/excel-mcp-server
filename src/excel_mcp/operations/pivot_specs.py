"""The values fields of a PivotTable: what they summarize and how their figures are shown."""

from dataclasses import replace

from excel_mcp.errors import InvalidArgumentError
from excel_mcp.operations.pivot_calc import CalcField
from excel_mcp.operations.pivot_fields import AxisField, FieldSetup, item_text
from excel_mcp.operations.pivot_options import PivotValue
from excel_mcp.operations.pivot_source import Source
from excel_mcp.operations.pivot_values import DataSpec

_CAPTIONS = {
    "sum": "Sum",
    "count": "Count",
    "average": "Average",
    "min": "Min",
    "max": "Max",
}
_NEEDS_BASE_FIELD = frozenset(
    {
        "percent_of_parent",
        "difference_from",
        "percent_difference_from",
        "percent_of",
        "running_total",
        "percent_running_total",
        "rank_ascending",
        "rank_descending",
    }
)
_COMPARES_WITH_ITEM = frozenset({"difference_from", "percent_difference_from", "percent_of"})


def data_specs(
    source: Source, setup: FieldSetup, calculated: list[CalcField], values: list[PivotValue]
) -> list[DataSpec]:
    specs: list[DataSpec] = []
    taken = {column.name.casefold() for column in source.columns}
    first_calculated = len(source.columns) + setup.extra_fields
    for value in values:
        key, field, name = _value_field(source, calculated, value, first_calculated)
        if key[0] == "calculated" and value.function != "sum":
            raise InvalidArgumentError(f"{name!r} is a calculated field, which can only be summed.")
        if key[0] == "column":
            column = source.columns[key[1]]
            if value.function != "count" and column.kind != "number":
                raise InvalidArgumentError(
                    f"Field {column.name!r} holds {column.kind}, not numbers, so it cannot be "
                    f"{value.function!r}-ed. Use function 'count'."
                )
        caption = _unique_caption(f"{_CAPTIONS[value.function]} of {name}", specs, taken)
        specs.append(DataSpec(key, field, value.function, caption, value.number_format))
    return specs


def _value_field(
    source: Source, calculated: list[CalcField], value: PivotValue, first_calculated: int
) -> tuple[tuple[str, int], int, str]:
    wanted = value.field.strip().casefold()
    for index, calc in enumerate(calculated):
        if calc.name.casefold() == wanted:
            return ("calculated", index), first_calculated + index, calc.name
    index = source.field_index(value.field)
    return ("column", index), index, source.columns[index].name


def _unique_caption(caption: str, specs: list[DataSpec], taken: set[str]) -> str:
    if caption.casefold() in taken:
        raise InvalidArgumentError(f"Rename the source column {caption!r}: it clashes.")
    used = {spec.caption.casefold() for spec in specs}
    candidate, number = caption, 1
    while candidate.casefold() in used:
        number += 1
        candidate = f"{caption}{number}"
    return candidate


def resolve_comparisons(
    specs: list[DataSpec], values: list[PivotValue], axis_fields: list[AxisField]
) -> list[DataSpec]:
    resolved = []
    for spec, value in zip(specs, values, strict=True):
        resolved.append(_resolve(spec, value, axis_fields))
    return resolved


def _resolve(spec: DataSpec, value: PivotValue, axis_fields: list[AxisField]) -> DataSpec:
    show_as = value.show_as
    if show_as is None:
        if value.base_field is not None or value.base_item is not None:
            raise InvalidArgumentError("base_field and base_item need show_as.")
        return spec
    if show_as not in _NEEDS_BASE_FIELD:
        if value.base_field is not None or value.base_item is not None:
            raise InvalidArgumentError(f"show_as {show_as!r} takes no base_field or base_item.")
        return replace(spec, show_as=show_as)
    if value.base_field is None:
        raise InvalidArgumentError(f"show_as {show_as!r} needs base_field.")
    wanted = value.base_field.strip().casefold()
    matches = [
        f for f in axis_fields if wanted in (f.name.casefold(), f.name.casefold().split(" (")[0])
    ]
    if not matches:
        names = [f.name for f in axis_fields]
        raise InvalidArgumentError(
            f"base_field {value.base_field!r} must be a field used in rows or columns: {names}."
        )
    base = matches[0]
    if base.levels > 2:
        raise InvalidArgumentError(
            f"base_field {value.base_field!r} is one of {base.levels} date groups; Excel's "
            "figures along three or more date levels are not supported. Group by fewer units."
        )
    if show_as == "percent_running_total" and base.hides_items:
        raise InvalidArgumentError(
            "show_as 'percent_running_total' needs all items of base_field shown: Excel "
            "includes the hidden ones in the total."
        )
    item = None
    if value.base_item is not None:
        if show_as not in _COMPARES_WITH_ITEM:
            raise InvalidArgumentError(f"show_as {show_as!r} takes no base_item.")
        texts = [item_text(label).casefold() for label in base.labels]
        if value.base_item.strip().casefold() not in texts:
            raise InvalidArgumentError(
                f"base_item {value.base_item!r} is not an item of {base.name!r}: "
                f"{[item_text(label) for label in base.labels][:20]}."
            )
        item = texts.index(value.base_item.strip().casefold())
    return replace(spec, show_as=show_as, base_field=base.index, base_item=item)


def sorted_by(field: AxisField, specs: list[DataSpec]) -> AxisField:
    if field.sort_by is None:
        return field
    wanted = field.sort_by.strip().casefold()
    for spec in specs:
        if spec.caption.casefold() == wanted:
            return replace(field, sort_by=spec.caption)
    raise InvalidArgumentError(
        f"sort_by {field.sort_by!r} must be a values field: {[spec.caption for spec in specs]}."
    )


def check_values_in_rows(specs: list[DataSpec], subtotals: bool, row_levels: int) -> None:
    """Excel mixes the values fields up in some figures when it lists them down the rows."""
    for spec in specs:
        if spec.show_as in ("rank_ascending", "rank_descending"):
            raise InvalidArgumentError(
                f"show_as {spec.show_as!r} ranks across all the values fields when values_in is "
                "'rows'. Use values_in 'columns'."
            )
        if (
            spec.show_as in _NEEDS_BASE_FIELD - {"percent_of_parent"}
            and subtotals
            and row_levels > 1
        ):
            raise InvalidArgumentError(
                f"show_as {spec.show_as!r} gives wrong subtotals when values_in is 'rows' with "
                "several row fields. Use subtotals=false or values_in 'columns'."
            )
