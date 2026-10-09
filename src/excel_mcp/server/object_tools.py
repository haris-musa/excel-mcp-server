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
        range: Annotated[RangeRef, Field(description="Including the header row, e.g. 'A1:D20'.")],
        options: Annotated[TableOptions, Field(default_factory=TableOptions)],
        name: Annotated[
            str | None, Field(description="Unique in the workbook. Default: TableN.")
        ] = None,
    ) -> Changed:
        """Turn a range with a header row of unique text labels into an Excel table.

        Options set the style, banding, header and totals rows, filter buttons, calculated
        columns and totals. Defaults are Excel's: striped rows, filter buttons.
        """
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
        name: Annotated[str, Field(description="Table name (describe_sheet lists them).")],
        options: Annotated[TableOptions, Field(default_factory=TableOptions)],
        range: Annotated[
            RangeRef | None,
            Field(
                description="Resize: the new range, with the same top-left cell. The totals row "
                "moves to the end. New columns take their header cell's text, or 'ColumnN'."
            ),
        ] = None,
    ) -> Changed:
        """Change a table's options, add calculated columns and totals, or resize it."""
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
                description="Kind of chart. waterfall, histogram, pareto, box_whisker, treemap, "
                "sunburst and funnel are Excel 2016 charts: they need `at` and take no "
                "combo, trendline, colors or secondary axis."
            ),
        ],
        options: Annotated[ChartOptions, Field(default_factory=ChartOptions)],
        at: Annotated[
            CellRef | None,
            Field(
                description="Top-left cell, e.g. 'E2'. Omit to put the chart on a new chart "
                "sheet named `sheet`."
            ),
        ] = None,
        source: Annotated[
            RangeRef | None,
            Field(
                description="A block with a header row, labels in the first column and one "
                "series per further column, e.g. 'A1:C13' or 'Data!A1:C13'. Scatter: x values "
                "first. Bubble: x, y, size. Excel 2016 charts: leading text columns are "
                "labels (the levels, for treemap and sunburst), the number columns the series."
            ),
        ] = None,
        series_in: Annotated[
            SeriesIn, Field(description="'rows': series are the rows of `source`.")
        ] = "columns",
        series: Annotated[
            list[SeriesSpec],
            Field(description="Explicit series instead of `source`, for any ranges on any sheet."),
        ] = [],  # noqa: B006
        categories: Annotated[
            str | None,
            Field(description="Category labels (x values) for series without their own."),
        ] = None,
        name: Annotated[
            str | None, Field(description="Unique among the sheet's charts and images.")
        ] = None,
        replace: Annotated[
            str | None,
            Field(
                description="Name of the chart to replace with this one (in the same position, "
                "keeping its name unless `name` is given), instead of adding a chart."
            ),
        ] = None,
    ) -> Changed:
        """Add a chart to `sheet`, or replace one. Returns its name and the cells it covers.

        Give `source` for a plain block, or `series` for ranges that are not adjacent, sit in
        rows, have their own names or live on other sheets. Combo charts take a `type` and
        `secondary_axis` per series. Options that do not fit the chart type are rejected.
        To chart a PivotTable, leave out its Grand Total row and column.
        describe_sheet lists charts; delete_chart removes one.
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
        name: Annotated[str, Field(description="Chart name from describe_sheet.")],
    ) -> Changed:
        """Remove a chart from a sheet. The data it plotted is left untouched."""
        with workspace.edit(path) as workbook:
            removed = chart_index.delete_chart(get_sheet(workbook, sheet), name)
        return Changed(sheet=sheet, name=removed.name, range=removed.range)

    @tools.writer("Add sparklines")
    def add_sparklines(
        path: WorkbookPath,
        sheet: SheetName,
        range: Annotated[
            RangeRef,
            Field(description="Cells that get a sparkline: one row or column, e.g. 'G2:G9'."),
        ],
        source: Annotated[
            RangeRef,
            Field(
                description="Data, on any sheet: 'B2:F9' or 'Data!B2:F9'. One sparkline per row."
            ),
        ],
        style: Annotated[SparklineStyle, Field(default_factory=SparklineStyle)],
    ) -> Changed:
        """Add a group of sparklines (Insert > Sparklines): a line, column or win/loss chart in
        each cell of `range`, one per row of `source` (or per column when the cell count
        matches the columns). Sparklines already in those cells are replaced.

        describe_sheet lists sparklines; delete_sparklines removes them.
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
        range: Annotated[RangeRef, Field(description="Cells whose sparklines to remove.")],
    ) -> Changed:
        """Remove the sparklines in a range (Clear Sparklines). The data is left untouched."""
        with workspace.edit(path) as workbook:
            sparklines.delete_sparklines(get_sheet(workbook, sheet), parse_range(range))
        return Changed(sheet=sheet, range=range)
