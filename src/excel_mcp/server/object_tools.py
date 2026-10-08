"""Tools for tables and charts."""

from typing import Annotated

from pydantic import Field

from excel_mcp.operations import chart_index, charts, tables
from excel_mcp.operations.charts_options import ChartOptions, ChartType
from excel_mcp.server.params import SheetName, WorkbookPath
from excel_mcp.server.registry import ToolRegistry
from excel_mcp.workspace import Workspace, get_sheet

DEFAULT_CHART_OPTIONS = ChartOptions()


def register(tools: ToolRegistry, workspace: Workspace) -> None:
    @tools.writer("Create table")
    def create_table(
        path: WorkbookPath,
        sheet: SheetName,
        range: Annotated[str, Field(description="Range including the header row, e.g. 'A1:D20'.")],
        name: Annotated[
            str | None, Field(description="Table name, unique in the workbook. Default: TableN.")
        ] = None,
        style: Annotated[
            str, Field(description="Excel table style, e.g. 'TableStyleMedium9'.")
        ] = "TableStyleMedium9",
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
        data_range: Annotated[
            str,
            Field(
                description="Data with a header row, labels in the first column and one "
                "series per further column, e.g. 'A1:C13'."
            ),
        ],
        chart_type: Annotated[ChartType, Field(description="Kind of chart to draw.")],
        anchor_cell: Annotated[str, Field(description="Cell where the chart's top-left sits.")],
        options: Annotated[
            ChartOptions,
            Field(
                description="Titles, size, legend, data labels, stacking, colors, markers, "
                "axis range and number format, and secondary-axis lines. Every field is "
                "optional."
            ),
        ] = DEFAULT_CHART_OPTIONS,
        data_sheet: Annotated[
            str | None, Field(description="Sheet holding the data. Default: `sheet`.")
        ] = None,
    ) -> str:
        """Add a chart that plots a block of data.

        'column' draws vertical bars, 'bar' horizontal bars. 'scatter' plots points, with the
        x values in the first column. 'doughnut' is a pie with a hole; 'radar' draws one
        polygon per series. Options that do not fit the chart type are rejected. List a
        sheet's charts with describe_sheet and remove one with delete_chart.
        """
        with workspace.edit(path) as workbook:
            area = charts.create_chart(
                get_sheet(workbook, sheet),
                get_sheet(workbook, data_sheet or sheet),
                data_range,
                chart_type,
                anchor_cell,
                options,
            )
        return f"Added a {chart_type} chart of {data_sheet or sheet}!{area} at {anchor_cell}."

    @tools.destroyer("Delete chart")
    def delete_chart(
        path: WorkbookPath,
        sheet: SheetName,
        index: Annotated[
            int,
            Field(description="1-based chart number, as listed under 'charts' by describe_sheet."),
        ],
    ) -> str:
        """Remove a chart from a sheet. The data it plotted is left untouched.

        Charts after the removed one move up by one index, so call describe_sheet again
        before deleting another.
        """
        with workspace.edit(path) as workbook:
            removed = chart_index.delete_chart(get_sheet(workbook, sheet), index)
        label = f" {removed.title!r}" if removed.title else ""
        return f"Deleted {removed.type} chart {index}{label} from {sheet}."
