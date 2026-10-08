<!-- mcp-name: io.github.haris-musa/excel-mcp-server -->
<p align="center">
  <img src="https://raw.githubusercontent.com/haris-musa/excel-mcp-server/main/assets/logo.png" alt="Excel MCP Server" width="300"/>
</p>

[![PyPI version](https://img.shields.io/pypi/v/excel-mcp-server.svg)](https://pypi.org/project/excel-mcp-server/)
[![Downloads](https://static.pepy.tech/badge/excel-mcp-server)](https://pepy.tech/project/excel-mcp-server)
[![CI](https://github.com/haris-musa/excel-mcp-server/actions/workflows/ci.yml/badge.svg)](https://github.com/haris-musa/excel-mcp-server/actions/workflows/ci.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)

A [Model Context Protocol](https://modelcontextprotocol.io) server that lets AI assistants
create, read and edit Excel workbooks. It needs no Microsoft Excel installation.

- **Read and write** cells, formulas (with results calculated for you) and dates, with paging and streaming reads for large sheets, and search
- **Format** fonts, fills, borders, number formats, column widths and frozen panes; hide or
  group rows, columns and sheets; set up printing; protect sheets
- **Structure** sheets (order, view, workbook settings and protection), rows and columns, merged cells, tables, charts (column, bar, line, area, pie, doughnut, radar, scatter and
  bubble, with combos, secondary axes, trendlines and error bars), images, hyperlinks and PivotTables
- **Data tools**: paste special, fill series, remove duplicates, text to columns, find and
  replace, sheet and table filters with criteria
- **Rules**: conditional formatting (scales, icon sets, top/bottom, duplicates, text, dates and
  more) and data validation (dropdowns from cells, limits, input messages, alert styles)
- **Macros**: read the VBA code in `.xlsm` files, module by module (never run). Writing VBA
  is off unless you start the server with `--allow-vba-write` (see below)
- **Safe by design**: optional folder confinement, a formula safety check, read-only mode,
  localhost-only HTTP by default, and atomic saves that never leave a half-written file

Works with `.xlsx`, `.xlsm` (macros are preserved), `.xltx` and `.xltm` files.

## Quick start

You need [uv](https://docs.astral.sh/uv/getting-started/installation/). Every client runs the
server with `uvx excel-mcp-server stdio`; replace `/path/to/workbooks` with the folder the
server may use.

**Claude Desktop (Chat)**: download `excel-mcp-server-<version>.mcpb` from the
[latest release](https://github.com/haris-musa/excel-mcp-server/releases/latest) and open it.
Claude asks which folder the server may use.

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
| `--max-file-mb N` | | `100` | Largest workbook the server opens |
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
| Objects | `create_table`, `create_chart`, `delete_chart`, `create_pivot_table`, `delete_pivot_table`, `insert_image`, `delete_image` |
| Names and notes | `set_defined_name`, `delete_defined_name`, `set_note`, `delete_note` |
| Macros | `read_vba`; with `--allow-vba-write`: `write_vba_module`, `delete_vba_module` |

Every parameter is documented in [TOOLS.md](TOOLS.md).

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

- Formulas are parsed before they are written. Functions that reach the network, other
  programs or host information (`WEBSERVICE`, `HYPERLINK`, `IMAGE`, `RTD`, `CALL`, `INFO`,
  `INDIRECT`, Google Sheets `IMPORTXML` and others), DDE links and references to other
  workbooks are rejected.
- Paths are resolved, including symlinks, before they are checked against the allowed folders.
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
  values fields as columns or rows. The figures written into the cells are the ones Excel shows
  after a refresh, checked against about 200 PivotTables recorded from real Excel
  (`tests/fixtures/pivot_golden.json`). The exceptions: items that tie when sorted by value may
  swap places on refresh, and `values_in: "rows"` cannot be combined with rank figures or, with
  subtotals, other figures along a field, because Excel mixes the values fields up there.
- Inserting or deleting rows and columns does not update formulas, charts or tables that
  refer to the moved cells.
- Workbook features openpyxl does not understand, such as shapes, slicers and some
  embedded objects, may be lost when a workbook is edited. Pictures, charts, tables and
  macros in `.xlsm` files are kept.

## Upgrading from 0.x

Version 1.0 renames and redesigns the tools, removes the SSE transport and requires
Python 3.11 or newer. The [changelog](CHANGELOG.md#100---2026-09-28) maps every old tool to
its replacement.

## Contributing

Contributions are welcome. See [CONTRIBUTING.md](CONTRIBUTING.md).

## Star history

[![Star History Chart](https://api.star-history.com/svg?repos=haris-musa/excel-mcp-server&type=Date)](https://www.star-history.com/#haris-musa/excel-mcp-server&Date)

## License

MIT. See [LICENSE](LICENSE).
