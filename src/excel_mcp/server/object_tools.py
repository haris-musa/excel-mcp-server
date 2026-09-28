"""Tools for tables, charts and summary tables."""

from typing import Annotated

from pydantic import Field

from excel_mcp.operations import charts, summary, tables
from excel_mcp.operations.charts import ChartOptions, ChartType
from excel_mcp.operations.summary import SummaryValue
from excel_mcp.server.params import CellRef, RangeRef, SheetName, WorkbookPath
from excel_mcp.server.registry import ToolRegistry
from excel_mcp.workspace import Workspace, get_sheet


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
            ChartOptions | None, Field(description="Titles, size and legend. Default: none.")
        ] = None,
        data_sheet: Annotated[
            str | None, Field(description="Sheet holding the data. Default: `sheet`.")
        ] = None,
    ) -> str:
        """Add a chart that plots a block of data.

        'column' draws vertical bars, 'bar' horizontal bars. For 'scatter', the first column
        holds the x values.
        """
        with workspace.edit(path) as workbook:
            area = charts.create_chart(
                get_sheet(workbook, sheet),
                get_sheet(workbook, data_sheet or sheet),
                data_range,
                chart_type,
                anchor_cell,
                options or ChartOptions(),
            )
        return f"Added a {chart_type} chart of {data_sheet or sheet}!{area} at {anchor_cell}."

    @tools.destroyer("Create summary table")
    def create_summary_table(
        path: WorkbookPath,
        sheet: SheetName,
        source_range: RangeRef,
        group_by: Annotated[
            list[str], Field(min_length=1, description="Header names to group by.")
        ],
        values: Annotated[
            list[SummaryValue], Field(min_length=1, description="Header names to aggregate.")
        ],
        target_sheet: SheetName,
        target_cell: CellRef = "A1",
    ) -> str:
        """Group rows and aggregate columns, like a pivot table, writing the result as cells.

        The result is a static table (not an Excel PivotTable) and does not update when the
        source changes. Only numeric values are summed, averaged or compared; 'count' counts
        non-empty cells. Formula cells are not evaluated.
        """
        with workspace.edit(path) as workbook:
            area = summary.create_summary(
                get_sheet(workbook, sheet),
                source_range,
                group_by,
                values,
                get_sheet(workbook, target_sheet),
                target_cell,
                workspace.limits.max_cells,
            )
        return f"Wrote the summary to {target_sheet}!{area}."
