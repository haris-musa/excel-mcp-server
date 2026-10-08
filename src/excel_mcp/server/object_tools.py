"""Tools for tables and charts."""

from typing import Annotated

from pydantic import Field

from excel_mcp.operations import chart_index, charts, tables
from excel_mcp.operations.charts_data import SeriesIn
from excel_mcp.operations.charts_options import ChartOptions, ChartType, SeriesSpec
from excel_mcp.server.params import SheetName, WorkbookPath
from excel_mcp.server.registry import ToolRegistry
from excel_mcp.workspace import Workspace, get_sheet


def register(tools: ToolRegistry, workspace: Workspace) -> None:
    @tools.writer("Create table")
    def create_table(
        path: WorkbookPath,
        sheet: SheetName,
        range: Annotated[str, Field(description="Including the header row, e.g. 'A1:D20'.")],
        name: Annotated[
            str | None, Field(description="Unique in the workbook. Default: TableN.")
        ] = None,
        style: Annotated[str, Field(description="Excel table style.")] = "TableStyleMedium9",
        striped_rows: Annotated[bool, Field(description="Shade alternate rows.")] = True,
    ) -> str:
        """Turn a range with a header row of unique text labels into an Excel table."""
        with workspace.edit(path) as workbook:
            table_name, area = tables.create_table(
                workbook, get_sheet(workbook, sheet), range, name, style, striped_rows
            )
        return f"Created table {table_name!r} at {sheet}!{area}."

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
                "first. Bubble: x, y, size. Excel 2016 charts: waterfall, pareto, funnel and "
                "box_whisker take labels then values; histogram, values only; treemap and "
                "sunburst, the hierarchy columns then the sizes."
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
