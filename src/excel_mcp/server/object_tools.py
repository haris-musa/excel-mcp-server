"""Tools for tables and charts."""

from typing import Annotated

from pydantic import Field

from excel_mcp.operations import chart_index, charts, sparklines, tables
from excel_mcp.operations.charts_data import SeriesIn
from excel_mcp.operations.charts_options import ChartOptions, ChartType, SeriesSpec
from excel_mcp.operations.sparkline_style import SparklineStyle
from excel_mcp.operations.table_options import TableOptions
from excel_mcp.refs import parse_range
from excel_mcp.server.params import SheetName, WorkbookPath
from excel_mcp.server.registry import ToolRegistry
from excel_mcp.workspace import Workspace, get_sheet


def register(tools: ToolRegistry, workspace: Workspace) -> None:
    @tools.writer("Create table")
    def create_table(
        path: WorkbookPath,
        sheet: SheetName,
        range: Annotated[str, Field(description="Including the header row, e.g. 'A1:D20'.")],
        options: Annotated[TableOptions, Field(default_factory=TableOptions)],
        name: Annotated[
            str | None, Field(description="Unique in the workbook. Default: TableN.")
        ] = None,
    ) -> str:
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
        return f"Created table {table_name!r} at {sheet}!{area}."

    @tools.destroyer("Edit table")
    def edit_table(
        path: WorkbookPath,
        sheet: SheetName,
        table: Annotated[str, Field(description="Table name (describe_sheet lists them).")],
        options: Annotated[TableOptions, Field(default_factory=TableOptions)],
        range: Annotated[
            str | None,
            Field(
                description="Resize: the new range, with the same top-left cell. Not with a "
                "totals row. New columns take their header cell's text, or 'ColumnN'."
            ),
        ] = None,
    ) -> str:
        """Change a table's options, add calculated columns and totals, or resize it."""
        with workspace.edit(path) as workbook:
            area = tables.edit_table(
                get_sheet(workbook, sheet), table, range, options, workspace.limits.max_cells
            )
        return f"Table {table!r} is now at {sheet}!{area}."

    @tools.writer("Create chart")
    def create_chart(
        path: WorkbookPath,
        sheet: SheetName,
        chart_type: Annotated[
            ChartType,
            Field(
                description="Kind of chart. waterfall, histogram, pareto, box_whisker, treemap, "
                "sunburst and funnel are Excel 2016 charts: they need anchor_cell and take no "
                "combo, trendline, colors or secondary axis."
            ),
        ],
        options: Annotated[ChartOptions, Field(default_factory=ChartOptions)],
        anchor_cell: Annotated[
            str | None,
            Field(
                description="Top-left cell, e.g. 'E2'. Omit to put the chart on a new chart "
                "sheet named `sheet`."
            ),
        ] = None,
        data_range: Annotated[
            str | None,
            Field(
                description="A block with a header row, labels in the first column and one "
                "series per further column, e.g. 'A1:C13' or 'Data!A1:C13'. Scatter: x values "
                "first. Bubble: x, y, size. Excel 2016 charts: leading text columns are "
                "labels (the levels, for treemap and sunburst), the number columns the series."
            ),
        ] = None,
        series_in: Annotated[
            SeriesIn, Field(description="'rows': series are the rows of data_range.")
        ] = "columns",
        series: Annotated[
            list[SeriesSpec],
            Field(
                description="Explicit series instead of data_range, for any ranges on any sheet."
            ),
        ] = [],  # noqa: B006
        categories: Annotated[
            str | None,
            Field(description="Category labels (x values) for series without their own."),
        ] = None,
        index: Annotated[
            int | None,
            Field(
                description="Replace chart N of `sheet` (from describe_sheet) with this one, "
                "rebuilt from these arguments, instead of adding a chart."
            ),
        ] = None,
    ) -> str:
        """Add a chart to `sheet`, or replace one.

        Give data_range for a plain block, or series for ranges that are not adjacent, sit in
        rows, have their own names or live on other sheets. Combo charts take a `type` and
        `secondary_axis` per series. To change a chart, create it again with `index`.
        Options that do not fit the chart type are rejected. describe_sheet lists the charts
        (those of Excel 2016 after the others); delete_chart removes one.
        """
        with workspace.edit(path) as workbook:
            return charts.create_chart(
                workbook,
                sheet,
                anchor_cell,
                chart_type,
                data_range,
                series_in,
                series,
                categories,
                options,
                index,
            )

    @tools.destroyer("Delete chart")
    def delete_chart(
        path: WorkbookPath,
        sheet: SheetName,
        index: Annotated[
            int,
            Field(description="Chart number from describe_sheet."),
        ],
    ) -> str:
        """Remove a chart from a sheet. The data it plotted is left untouched.

        Later charts move up one index; call describe_sheet again before deleting another.
        """
        with workspace.edit(path) as workbook:
            removed = chart_index.delete_chart(get_sheet(workbook, sheet), index)
        label = f" {removed.title!r}" if removed.title else ""
        return f"Deleted {removed.type} chart {index}{label} from {sheet}."

    @tools.writer("Add sparklines")
    def add_sparklines(
        path: WorkbookPath,
        sheet: SheetName,
        location: Annotated[
            str, Field(description="Cells that get a sparkline: one row or column, e.g. 'G2:G9'.")
        ],
        data: Annotated[
            str,
            Field(
                description="Data, on any sheet: 'B2:F9' or 'Data!B2:F9'. One sparkline per row."
            ),
        ],
        style: Annotated[SparklineStyle, Field(default_factory=SparklineStyle)],
    ) -> str:
        """Add a group of sparklines (Insert > Sparklines): a line, column or win/loss chart in
        each cell of `location`, one per row of `data` (or per column when the cell count
        matches the columns). Sparklines already in those cells are replaced.

        describe_sheet lists sparklines; delete_sparklines removes them.
        """
        with workspace.edit(path) as workbook:
            return sparklines.add_sparklines(
                get_sheet(workbook, sheet), location, data, style, workspace.limits.max_cells
            )

    @tools.destroyer("Delete sparklines")
    def delete_sparklines(
        path: WorkbookPath,
        sheet: SheetName,
        range: Annotated[str, Field(description="Cells whose sparklines to remove.")],
    ) -> str:
        """Remove the sparklines in a range (Clear Sparklines). The data is left untouched."""
        with workspace.edit(path) as workbook:
            removed = sparklines.delete_sparklines(get_sheet(workbook, sheet), parse_range(range))
        return f"Deleted {removed} sparklines from {sheet}!{range}."
