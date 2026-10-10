<!-- mcp-name: io.github.haris-musa/excel-mcp-server -->
<p align="center">
  <picture>
    <source media="(prefers-color-scheme: dark)" srcset="https://raw.githubusercontent.com/haris-musa/excel-mcp-server/main/assets/logo-dark.png">
    <img src="https://raw.githubusercontent.com/haris-musa/excel-mcp-server/main/assets/logo.png" alt="Excel MCP Server" width="420">
  </picture>
</p>

<p align="center">
  <a href="https://pypi.org/project/excel-mcp-server/"><img src="https://img.shields.io/pypi/v/excel-mcp-server.svg" alt="PyPI version"></a>
  <a href="https://pypi.org/project/excel-mcp-server/"><img src="https://img.shields.io/pypi/pyversions/excel-mcp-server.svg" alt="Python versions"></a>
  <a href="https://pepy.tech/project/excel-mcp-server"><img src="https://static.pepy.tech/badge/excel-mcp-server" alt="Downloads"></a>
  <a href="https://github.com/haris-musa/excel-mcp-server/actions/workflows/ci.yml"><img src="https://github.com/haris-musa/excel-mcp-server/actions/workflows/ci.yml/badge.svg" alt="CI"></a>
  <a href="LICENSE"><img src="https://img.shields.io/badge/License-MIT-yellow.svg" alt="License: MIT"></a>
</p>

<p align="center">
  <a href="https://trendshift.io/repositories/13975" target="_blank" rel="noopener noreferrer"><img src="https://trendshift.io/api/badge/repositories/13975" alt="haris-musa%2Fexcel-mcp-server | Trendshift" width="250" height="55"/></a>
</p>

# Excel MCP Server

An open-source [Model Context Protocol](https://modelcontextprotocol.io) (MCP) server that lets
Claude, Cursor, VS Code Copilot, Codex and other AI clients create, read and edit Excel files:
charts, PivotTables, formulas, tables and more. It is built on `openpyxl`, needs **no Microsoft
Excel**, and runs on Windows, macOS and Linux.

**Install and try it at [excelmcpserver.com](https://excelmcpserver.com)**, or see
[Quick start](#quick-start) below.

## Features

**Charts and PivotTables**
- Create Excel charts with AI: column, bar, line, area, pie, doughnut, radar, scatter and
  bubble, with combos, secondary axes, trendlines and error bars
- The Excel 2016 chart types: waterfall, histogram, Pareto, box and whisker, treemap, sunburst
  and funnel
- Real PivotTables that Excel refreshes (grouping, calculated fields, "show values as",
  layouts, filters), plus slicers and timelines for PivotTables and tables
- Sparklines (line, column, win/loss)

**Formulas and dynamic arrays**
- Formulas are calculated without Excel by a built-in engine (about 260 functions, checked
  against more than 2,500 formulas recorded from real Excel), so `read_range` returns values
  for files that were never opened in Excel
- Dynamic arrays (`FILTER`, `SORT`, `UNIQUE`, `SEQUENCE`, `LET`) are stored as Excel stores
  them and spill into the cells next to them

**Data, tables and rules**
- Read and write cells, dates and hyperlinks; tables with totals rows and calculated columns
- Conditional formatting (scales, data bars, icon sets, top/bottom, duplicates, text, dates),
  data validation (dropdowns, limits, input messages)
- Sort, filter (with criteria), find and replace, remove duplicates, text to columns, fill
  series, paste special
- Fonts, fills, borders, number formats, merged cells, frozen panes, grouping, print setup,
  sheet protection, images, notes and defined names. Inserting or deleting rows and columns
  updates every reference, as in Excel

**Large files**
- Reads stream the workbook and return pages, so memory stays flat on sheets with hundreds of
  thousands of rows; every call is bounded by size and cell limits

**Macros (opt-in)**
- Read the VBA code in `.xlsm` files. Writing VBA is off unless you start the server with
  `--allow-vba-write`; the server never runs macros

**Safe by design**
- Optional folder confinement, a formula safety check, read-only mode, localhost-only HTTP by
  default, and atomic saves that never leave a half-written file

Works with `.xlsx`, `.xlsm` (macros are preserved), `.xltx` and `.xltm` files. Existing
threaded comments, shapes, form controls and other content the tools cannot edit are kept as
they are.

## Quick start

You need [uv](https://docs.astral.sh/uv/getting-started/installation/). Every client runs the
server with `uvx excel-mcp-server stdio`; replace `/path/to/workbooks` with the folder the
server may use. [excelmcpserver.com](https://excelmcpserver.com) has the same steps with copy
buttons and one-click install links for Cursor and VS Code.

**Claude Desktop (Chat)**: download
[`excel-mcp-server.mcpb`](https://github.com/haris-musa/excel-mcp-server/releases/latest/download/excel-mcp-server.mcpb)
from the latest release and open it. Claude asks which folder the server may use.

**Claude Code** (the CLI, and the Code tab in Claude Desktop):

```bash
claude mcp add excel --scope user -- uvx excel-mcp-server stdio --allow-dir /path/to/workbooks
```

**Cursor** (`~/.cursor/mcp.json`), and most clients that use an `mcpServers` config:

```json
{
  "mcpServers": {
    "excel": {
      "command": "uvx",
      "args": ["excel-mcp-server", "stdio", "--allow-dir", "/path/to/workbooks"]
    }
  }
}
```

**VS Code with GitHub Copilot** (`.vscode/mcp.json`, note the `servers` key):

```json
{
  "servers": {
    "excel": {
      "type": "stdio",
      "command": "uvx",
      "args": ["excel-mcp-server", "stdio", "--allow-dir", "${workspaceFolder}"]
    }
  }
}
```

<details>
<summary>OpenAI Codex, Gemini CLI, Devin Desktop, and Claude Desktop without the bundle</summary>

**OpenAI Codex** (CLI, IDE extension and app):

```bash
codex mcp add excel -- uvx excel-mcp-server stdio --allow-dir /path/to/workbooks
```

**Gemini CLI**:

```bash
gemini mcp add -s user excel uvx excel-mcp-server stdio --allow-dir /path/to/workbooks
```

**Devin Desktop**:

```bash
devin mcp add -s user excel -- uvx excel-mcp-server stdio --allow-dir /path/to/workbooks
```

**Claude Desktop, manual setup**: open Settings, Developer, Edit Config, add the `mcpServers`
JSON above to `claude_desktop_config.json`, and restart Claude.

</details>

Desktop apps often cannot find `uvx`, because they do not see your shell's `PATH` (common
on macOS). Use its full path instead, which `which uvx` prints.

## Choosing which files it can use

With `--allow-dir DIR`, the server only opens workbooks inside `DIR` (subfolders included)
and relative paths such as `reports/q1.xlsx` start there. Repeat the flag to allow several
folders, or set `EXCEL_FILES_PATH` (separate folders with `:` on macOS/Linux and `;` on
Windows).

Without `--allow-dir`, any absolute path to an Excel file works. In every mode the server
only touches Excel files and never silently overwrites an existing workbook.

## Remote use (Streamable HTTP)

```bash
uvx excel-mcp-server streamable-http --allow-dir /srv/workbooks
```

Clients connect to `http://127.0.0.1:8017/mcp`. Workbooks live in the `--allow-dir` folder
(default `./excel_files`), and `export_workbook` / `import_workbook` move files between the
server and the client.

The server listens on localhost only. To accept other machines, set a token and a host:

```bash
EXCEL_MCP_AUTH_TOKEN=change-me uvx excel-mcp-server streamable-http --host 0.0.0.0
```

Clients then send `Authorization: Bearer change-me`. Put a TLS-terminating reverse proxy in
front of it for use across networks.

### Docker

```bash
docker build -t excel-mcp-server .
docker run -p 8017:8017 -v "$PWD/workbooks:/data" -e EXCEL_MCP_AUTH_TOKEN=change-me excel-mcp-server
```

## Configuration

| Flag | Environment variable | Default | Meaning |
| --- | --- | --- | --- |
| `--allow-dir DIR` | `EXCEL_FILES_PATH` | none (stdio), `./excel_files` (HTTP) | Folders workbooks must be in |
| `--read-only` | `EXCEL_MCP_READ_ONLY=1` | off | Only offer tools that do not change files |
| `--allow-vba-write` | `EXCEL_MCP_ALLOW_VBA_WRITE=1` | off | Add tools that write VBA macros (see warning below); ignored with `--read-only` |
| `--max-file-mb N` | | `100` | Largest workbook the server opens. Also sets the unpacked-size limit (10 times this) that every opened or uploaded workbook must stay under; parts compressed over 500:1 are refused |
| `--log-level LEVEL` | | `WARNING` | Logging on stderr |
| `--host HOST` | `EXCEL_MCP_HOST` | `127.0.0.1` | HTTP listen address |
| `--port PORT` | `EXCEL_MCP_PORT` | `8017` | HTTP port |
| | `EXCEL_MCP_AUTH_TOKEN` | none | Bearer token required on HTTP requests |
| `--allow-unauthenticated` | | off | Allow a non-local `--host` without a token |

## Tools

| Area | Tools |
| --- | --- |
| Workbooks | `create_workbook`, `describe_workbook`, `set_workbook_settings`, `list_workbooks`, `export_workbook`, `import_workbook` |
| Sheets | `describe_sheet`, `create_sheet`, `rename_sheet`, `copy_sheet`, `delete_sheet`, `insert_rows_or_columns`, `delete_rows_or_columns` |
| Cells | `read_range`, `write_range`, `clear_range`, `copy_range`, `sort_range`, `transform_range`, `find_cells`, `replace_cells` |
| Formatting | `format_range`, `merge_cells`, `set_sheet_layout`, `add_conditional_format`, `add_data_validation` |
| Objects | `set_table`, `create_chart`, `delete_chart`, `create_pivot_table`, `delete_pivot_table`, `add_slicer`, `delete_slicer`, `insert_image`, `delete_image`, `add_sparklines`, `delete_sparklines` |
| Names and notes | `set_defined_name`, `delete_defined_name`, `set_note`, `delete_note` |
| Macros | `read_vba`; with `--allow-vba-write`: `write_vba_module`, `delete_vba_module` |

Every parameter is documented in [TOOLS.md](TOOLS.md). Parameters mean the same in every tool:

| Parameter | Meaning |
| --- | --- |
| `path`, `sheet` | The workbook and the worksheet |
| `range`, `at` | A block of cells (`A1:D20`), and the top-left cell of whatever is placed there: written values, copies, charts, images, PivotTables, slicers |
| `source` | Where data comes from, optionally on another sheet (`Data!A1:E200`) |
| `name` | The name of a table, chart, image, PivotTable, slicer or defined name |
| `_cm`, `_pt`, `_chars` | Units of sizes: `width_cm`, `row_heights_pt`, `column_widths_chars` |

Tools that change a workbook return what they changed (`sheet`, `range`, `name`) so that the
next call can refer to it. Charts and images are selected by name, as listed by `describe_sheet`.

### Writing macros (opt-in)

With `--allow-vba-write` (or `EXCEL_MCP_ALLOW_VBA_WRITE=1`) the server can set the code of
standard, class, workbook and sheet modules in `.xlsm` and `.xltm` files, delete standard and
class modules, and create new `.xlsm`/`.xltm` workbooks. The server only stores the code; it
never runs it, and Excel asks before enabling macros. Digitally signed VBA projects are
refused, since any change would invalidate the signature.

> **Warning:** macro code runs with your user's rights once you enable macros in Excel. A
> model that reads untrusted content could be tricked into writing harmful code. Enable this
> option only for trusted workflows, keep `--allow-dir` narrow, and **read the code before
> you enable macros**. The tools are never offered in `--read-only` mode.

## Security

- Formulas are parsed before they are written. Functions that reach the network or other
  programs (`WEBSERVICE`, `IMAGE`, `RTD`, `CALL`, Google Sheets `IMPORTXML` and others),
  Excel 4.0 macro functions (`FILES`, `GET.WORKBOOK`, `RUN` and others), DDE links and
  references to other workbooks are rejected. `HYPERLINK` is accepted with a fixed
  `http(s)://`, `mailto:` or `#Sheet!A1` link only.
- Paths are resolved, including symlinks, before they are checked against the allowed
  folders. Network (UNC) and Windows device paths are always rejected.
- A single call processes at most 100,000 cells, and large reads are returned in pages.
- Cell contents are data from files. The server tells the model not to follow instructions
  found in them, but review what an assistant does with workbooks from untrusted sources.

Please report vulnerabilities privately; see [SECURITY.md](SECURITY.md).

## Limitations

- Formulas are stored in the file, and Excel calculates them when it opens it. So that
  `read_range` is useful before that, `values` mode calculates formulas that have no saved
  result with a built-in calculator (about 260 functions: math, statistics, financial and
  bond, dates, text, lookup, logical, LET, sorting and filtering). Results Excel saved are
  always preferred. A formula the calculator cannot reproduce exactly as Excel does (an
  unsupported function, a circular reference, or an Excel quirk it does not replicate) is
  returned as null and listed in `uncalculated` with the reason; it is never guessed.
  Ranges in a formula are reduced to the formula's own row or column where Excel's ordinary
  (not array-entered) formulas do the same, and volatile functions such as `RAND` are not
  calculated. The calculator is checked against more than 2,500 formulas recorded from real
  Excel (`tests/fixtures/formula_golden.json`).
- Functions Excel added after 2007 (`IFS`, `XLOOKUP`, `STDEV.S`, `SORT`, ...) are written
  with the `_xlfn.` prefix Excel expects, wherever formulas are stored (cells, copied and
  sorted cells, conditional formats, data validation, defined names).
- Legacy `.xls` and `.csv` files are not supported.
- PivotTables are created from a snapshot and can use text, number and date columns of up to
  100,000 cells; Excel refreshes them from the live source data. They support number formats,
  "show values as" (percent of total, row, column or parent, difference from, running total,
  rank), sorting by label or value, date and number grouping, calculated fields, compact,
  outline and tabular layouts, subtotals on or off, filters with chosen items and several
  value fields as columns or rows. The figures written into the cells are the ones Excel shows
  after a refresh, checked against about 200 PivotTables recorded from real Excel
  (`tests/fixtures/pivot_golden.json`). The exceptions: items that tie when sorted by value may
  swap places on refresh, and `values_in: "rows"` cannot be combined with rank figures or, with
  subtotals, other figures along a field, because Excel mixes the value fields up there.
- `add_slicer` adds slicers (Insert > Slicer) for tables and PivotTables, and timelines for date
  fields of PivotTables: field, caption, position and size, columns, style, sort, selected
  items (or a period) and "hide items with no data". Selecting items filters as clicking
  does: table rows are filtered and hidden; PivotTable items are hidden and the figures
  recalculated from the source, which Excel confirms on refresh (about 200 PivotTables are
  checked against Excel, and the slicer variants against Excel's own results). One slicer can
  filter several PivotTables that share a cache, such as those of copied sheets. Limits: a
  PivotTable made or refreshed by Excel keeps its figures in the cells (the hidden items and
  a refresh-on-open mark are stored, so Excel recalculates when it opens the file); fields
  grouped in a PivotTable take no slicer.
- Inserting or deleting rows and columns, and renaming a sheet, update references like
  Excel. Hyperlink targets and 3D references (`Sheet1:Sheet3!A1`) are left alone (Excel does
  the same). An edit that cuts through an array formula, a PivotTable, a table header or two
  tables at once is refused. Inserted cells take the formatting, row height and column width
  of the line above or to the left (nothing at the first row or column), and calculated table
  columns are filled into rows inserted in a table.
- Editing a workbook keeps what the tools cannot change: sparklines, extended conditional
  formats and validation, slicers and timelines, newer charts (waterfall, histogram,
  treemap and others), threaded comments, shapes, form controls, linked data types and
  custom XML stay as Excel saved them and move with inserted or deleted rows and columns and
  renamed sheets. `copy_sheet` copies sparklines, slicers and timelines (as Excel does), the Excel
  2010 half of conditional formats (data bars, icon sets) and the Excel 2016 charts made by
  `create_chart`, but not the others. Deleting a sheet removes slicers that only it used;
  slicers that would be left without their PivotTable block the deletion, and deleting a
  table, or the column a table slicer filters, removes the slicer.
  Digital signatures are removed, as Excel does when a signed file changes.
- Formulas that return several values (`FILTER`, `SORT`, `UNIQUE`, `SEQUENCE`, `A2:A9*2`) are
  stored as dynamic array formulas, as Excel stores them, and spill into the cells next to
  them. A cell in the way blocks the spill (Excel shows `#SPILL!`), and `write_range`
  reports it. Where the calculator cannot tell how large a result is, only the formula's
  cell is stored and Excel fills in the rest when it opens the file.

## Upgrading from 1.x

Version 2.0 makes tool and parameter names consistent across tools, which breaks 1.x clients
and saved prompts. The changelog has a table that maps every 1.x name to its 2.0 replacement:
[Upgrading from 1.x](CHANGELOG.md#upgrading-from-1x). Coming from 0.x? See the
[1.0.0 notes](CHANGELOG.md#100---2026-09-28) first.

## Contributing

Contributions are welcome. See [CONTRIBUTING.md](CONTRIBUTING.md).

## License

MIT. See [LICENSE](LICENSE).
