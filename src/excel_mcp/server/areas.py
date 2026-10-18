"""Tool areas: groups of tools that can be switched on together with ``--tools``."""

# Offered whatever the areas, so a client can always look at a workbook.
CORE_TOOLS = ("describe_workbook", "describe_sheet", "read_range")

AREAS: dict[str, tuple[str, ...]] = {
    "workbooks": (
        "create_workbook",
        "set_workbook_settings",
        "list_workbooks",
        "export_workbook",
        "import_workbook",
    ),
    "csv": ("import_csv", "export_csv"),
    "sheets": (
        "create_sheet",
        "rename_sheet",
        "copy_sheet",
        "delete_sheet",
        "insert_rows_or_columns",
        "delete_rows_or_columns",
    ),
    "data": (
        "write_range",
        "clear_range",
        "copy_range",
        "sort_range",
        "transform_range",
        "find_cells",
        "replace_cells",
    ),
    "formulas": ("trace_formula", "goal_seek"),
    "formatting": (
        "format_range",
        "merge_cells",
        "set_sheet_layout",
        "add_conditional_format",
        "add_data_validation",
    ),
    "tables": ("set_table",),
    "charts": ("create_chart", "delete_chart", "add_sparklines", "delete_sparklines"),
    "pivots": ("create_pivot_table", "delete_pivot_table"),
    "slicers": ("add_slicer", "delete_slicer"),
    "names": ("set_defined_name", "delete_defined_name"),
    "notes": ("set_note", "delete_note"),
    "images": ("insert_image", "delete_image"),
    "vba": ("read_vba", "write_vba_module", "delete_vba_module"),
}


def parse_areas(text: str) -> tuple[str, ...]:
    """The area names in a comma-separated list, or a ValueError naming the valid ones."""
    names = tuple(dict.fromkeys(name.strip().lower() for name in text.split(",") if name.strip()))
    unknown = [name for name in names if name not in AREAS]
    if unknown or not names:
        problem = f"unknown tool area {', '.join(unknown)}" if unknown else "no tool area given"
        raise ValueError(f"{problem}; choose from {', '.join(AREAS)}")
    return names


def tool_names(areas: tuple[str, ...] | None) -> frozenset[str] | None:
    """The tools to offer: the core and the chosen areas; None means every tool."""
    if areas is None:
        return None
    return frozenset(CORE_TOOLS).union(*(AREAS[area] for area in areas))
