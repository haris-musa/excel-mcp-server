"""Tools for tables, charts and sparklines."""

from typing import Annotated

from pydantic import Field

from excel_mcp.operations import chart_index, charts, sparklines, tables
from excel_mcp.operations.charts_data import SeriesIn
from excel_mcp.operations.charts_options import ChartOptions, ChartType, SeriesSpec
from excel_mcp.operations.sparkline_style import SparklineStyle
from excel_mcp.operations.table_options import TableOptions
from excel_mcp.refs import parse_range
from excel_mcp.server.params import CellRef, RangeRef, SheetName, WorkbookPath
from excel_mcp.server.registry import ToolRegistry
from excel_mcp.server.results import Changed
from excel_mcp.workspace import Workspace, get_sheet


def register(tools: ToolRegistry, workspace: Workspace) -> None:
    @tools.writer("Create table")
    def create_table(
        path: WorkbookPath,
        sheet: SheetName,
        range: Annotated[RangeRef, Field(description="With the header row.")],
        options: Annotated[TableOptions, Field(default_factory=TableOptions)],
        name: Annotated[str | None, Field(description="Default: TableN.")] = None,
    ) -> Changed:
        """Turn a range with a header row of unique text labels into a table, as Excel does."""
        with workspace.edit(path) as workbook:
            table_name, area = tables.create_table(
                workbook,
                get_sheet(workbook, sheet),
                range,
                name,
                options,
                workspace.limits.max_cells,
            )
        return Changed(sheet=sheet, name=table_name, range=area)

    @tools.destroyer("Edit table")
    def edit_table(
        path: WorkbookPath,
        sheet: SheetName,
        name: str,
        options: Annotated[TableOptions, Field(default_factory=TableOptions)],
        range: Annotated[
            RangeRef | None,
            Field(description="Resize to this range (same top-left cell)."),
        ] = None,
    ) -> Changed:
        """Change a table's options, calculated columns and totals, or resize it."""
        with workspace.edit(path) as workbook:
            area = tables.edit_table(
                get_sheet(workbook, sheet), name, range, options, workspace.limits.max_cells
            )
        return Changed(sheet=sheet, name=name, range=area)

    @tools.writer("Create chart")
    def create_chart(
        path: WorkbookPath,
        sheet: SheetName,
        chart_type: Annotated[
            ChartType,
            Field(
                description="Excel 2016 types (waterfall to funnel) need `at`; they take no "
                "combo, trendline, colors or secondary axis."
            ),
        ],
        options: Annotated[ChartOptions, Field(default_factory=ChartOptions)],
        at: Annotated[
            CellRef | None,
            Field(description="Top-left cell. Omit for a new chart sheet named `sheet`."),
        ] = None,
        source: Annotated[
            RangeRef | None,
            Field(
                description="Header row, labels in the first column, a series per further "
                "column, e.g. 'A1:C13'. Scatter: x first. Bubble: x, y, size. Excel 2016 "
                "types: leading text columns are labels (levels), number columns the series."
            ),
        ] = None,
        series_in: Annotated[
            SeriesIn, Field(description="'rows': one series per row of `source`.")
        ] = "columns",
        series: Annotated[
            list[SeriesSpec],
            Field(description="Instead of `source`: any ranges, on any sheet."),
        ] = [],  # noqa: B006
        categories: Annotated[
            str | None,
            Field(description="Labels for series without their own."),
        ] = None,
        name: Annotated[
            str | None, Field(description="Unique among the sheet's charts and images.")
        ] = None,
        replace: Annotated[
            str | None,
            Field(
                description="Name of a chart to replace in place, keeping its name unless "
                "`name` is given."
            ),
        ] = None,
    ) -> Changed:
        """Add or replace a chart. Returns its name and cells.

        `source` is a plain block; `series` is for non-adjacent ranges, rows, names or other
        sheets, and combos (`type`, `secondary_axis` per series). Options that do not fit the
        chart type are rejected. For a PivotTable, leave out the Grand Total row and column.
        """
        with workspace.edit(path) as workbook:
            made, area = charts.create_chart(
                workbook,
                sheet,
                at,
                chart_type,
                source,
                series_in,
                series,
                categories,
                options,
                replace,
                name,
            )
        return Changed(sheet=sheet, name=made, range=area)

    @tools.destroyer("Delete chart")
    def delete_chart(
        path: WorkbookPath,
        sheet: SheetName,
        name: Annotated[str, Field(description="Name from describe_sheet.")],
    ) -> Changed:
        """Remove a chart; its data stays."""
        with workspace.edit(path) as workbook:
            removed = chart_index.delete_chart(get_sheet(workbook, sheet), name)
        return Changed(sheet=sheet, name=removed.name, range=removed.range)

    @tools.writer("Add sparklines")
    def add_sparklines(
        path: WorkbookPath,
        sheet: SheetName,
        range: Annotated[
            RangeRef,
            Field(description="One row or column, e.g. 'G2:G9'."),
        ],
        source: Annotated[
            RangeRef,
            Field(description="Data, e.g. 'Data!B2:F9': a sparkline per row."),
        ],
        style: Annotated[SparklineStyle, Field(default_factory=SparklineStyle)],
    ) -> Changed:
        """Add sparklines (Insert > Sparklines) to the cells of `range`, one per row of
        `source` (per column when the cells match the columns). Replaces existing ones there.
        """
        with workspace.edit(path) as workbook:
            area, replaced = sparklines.add_sparklines(
                get_sheet(workbook, sheet), range, source, style, workspace.limits.max_cells
            )
        note = f"Replaced {replaced} existing sparklines." if replaced else None
        return Changed(sheet=sheet, range=str(area), note=note)

    @tools.destroyer("Delete sparklines")
    def delete_sparklines(
        path: WorkbookPath,
        sheet: SheetName,
        range: RangeRef,
    ) -> Changed:
        """Remove the sparklines in a range."""
        with workspace.edit(path) as workbook:
            sparklines.delete_sparklines(get_sheet(workbook, sheet), parse_range(range))
        return Changed(sheet=sheet, range=range)
