"""Tools for CSV files: import into a sheet and export a range."""

from typing import Annotated, Literal

from pydantic import Field

from excel_mcp.operations.calculated import MAX_LISTED
from excel_mcp.operations.cells import WriteResult
from excel_mcp.operations.csv_export import export_csv as render_csv
from excel_mcp.operations.csv_format import Delimiter, Encoding
from excel_mcp.operations.csv_import import FieldRules
from excel_mcp.operations.csv_import import import_csv as read_csv
from excel_mcp.operations.csv_values import DateOrder
from excel_mcp.operations.sheets import validate_sheet_name
from excel_mcp.server.params import CellRef, RangeRef, SheetName, WorkbookPath
from excel_mcp.server.registry import ToolRegistry
from excel_mcp.server.results import Changed
from excel_mcp.workspace import Workspace, get_sheet

CsvPath = Annotated[str, Field(description=".csv or .tsv file, located like workbooks.")]


def register(tools: ToolRegistry, workspace: Workspace) -> None:
    limits = workspace.limits

    @tools.destroyer("Import CSV")
    def import_csv(
        path: WorkbookPath,
        csv_path: CsvPath,
        sheet: Annotated[SheetName, Field(description="Created if missing.")],
        at: Annotated[CellRef, Field(description="Top-left cell.")] = "A1",
        delimiter: Annotated[
            Literal["auto", "comma", "semicolon", "tab", "pipe"],
            Field(description="'auto': most frequent in the first record."),
        ] = "auto",
        encoding: Annotated[
            Encoding,
            Field(description="'utf-8' also reads a BOM."),
        ] = "utf-8",
        decimal: Annotated[
            Literal[".", ","], Field(description="Decimal mark; the other groups thousands.")
        ] = ".",
        date_order: Annotated[
            DateOrder,
            Field(
                description="'dmy' or 'mdy' also read 31/12/2026-style dates; "
                "else only ISO ones are dates."
            ),
        ] = "ymd",
        as_text: Annotated[bool, Field(description="Keep every field as text.")] = False,
    ) -> WriteResult:
        """Import a CSV file into a sheet from `at`, overwriting cells.

        Numbers, TRUE/FALSE and ISO dates become values; leading zeros ('007'), over 15 digits
        and ambiguous dates stay text. Never formulas: text starting with = + - @ is text.
        """
        file = workspace.resolve_csv_existing(csv_path)
        rules = FieldRules(decimal=decimal, date_order=date_order, as_text=as_text)
        with workspace.edit(path) as workbook:
            if sheet not in workbook.sheetnames:
                validate_sheet_name(sheet, workbook.sheetnames)
                workbook.create_sheet(sheet)
            return read_csv(
                get_sheet(workbook, sheet),
                file,
                at,
                delimiter=delimiter,
                encoding=encoding,
                rules=rules,
                max_cells=limits.max_cells,
            )

    @tools.destroyer("Export CSV")
    def export_csv(
        path: WorkbookPath,
        sheet: SheetName,
        csv_path: CsvPath,
        range: Annotated[
            RangeRef | None, Field(description="e.g. 'A1:D20'. Default: the used range.")
        ] = None,
        formatted: Annotated[
            bool,
            Field(description="Apply number formats, as Excel shows cells; else raw values."),
        ] = False,
        delimiter: Delimiter = "comma",
        encoding: Annotated[
            Encoding, Field(description="'utf-8-sig' adds a BOM, which Excel needs for UTF-8.")
        ] = "utf-8",
        escape_formulas: Annotated[
            bool,
            Field(description="Prefix text starting with = + - @ with ' so Excel does not run it."),
        ] = True,
        overwrite: Annotated[
            bool, Field(description="Replace an existing file (else an error).")
        ] = False,
    ) -> Changed:
        """Export a range (formula results) as a CSV file, like Save As CSV.

        Text starting with = + - @ gets a ' prefix unless `escape_formulas` is false.
        """
        result = render_csv(
            workspace,
            path,
            sheet,
            range,
            formatted=formatted,
            delimiter=delimiter,
            encoding=encoding,
            escape_formulas=escape_formulas,
            max_cells=limits.max_cells,
        )
        stored = workspace.store_csv(csv_path, result.content, overwrite=overwrite)
        note = None
        if result.uncalculated:
            listed = ", ".join(list(result.uncalculated)[:MAX_LISTED])
            note = f"Formulas without a result were written empty: {listed}."
        return Changed(path=workspace.display(stored), sheet=sheet, range=result.range, note=note)
