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

- **Read and write** cells, formulas and dates, with paging for large sheets and search
- **Format** fonts, fills, borders, number formats, column widths and frozen panes
- **Structure** sheets, rows and columns, merged cells, tables, charts and summary tables
- **Rules**: conditional formatting and data validation (dropdowns, number limits)
- **Macros**: read the VBA code in `.xlsm` files, module by module (read-only, never run)
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
| `--max-file-mb N` | | `100` | Largest workbook the server opens |
| `--log-level LEVEL` | | `WARNING` | Logging on stderr |
| `--host HOST` | `EXCEL_MCP_HOST` | `127.0.0.1` | HTTP listen address |
| `--port PORT` | `EXCEL_MCP_PORT` | `8017` | HTTP port |
| | `EXCEL_MCP_AUTH_TOKEN` | none | Bearer token required on HTTP requests |
| `--allow-unauthenticated` | | off | Allow a non-local `--host` without a token |

## Tools

| Area | Tools |
| --- | --- |
| Workbooks | `create_workbook`, `describe_workbook`, `list_workbooks`, `export_workbook`, `import_workbook` |
| Sheets | `describe_sheet`, `create_sheet`, `rename_sheet`, `copy_sheet`, `delete_sheet`, `insert_rows_or_columns`, `delete_rows_or_columns` |
| Cells | `read_range`, `write_range`, `clear_range`, `copy_range`, `find_cells` |
| Formatting | `format_range`, `merge_cells`, `set_sheet_layout`, `add_conditional_format`, `add_data_validation` |
| Objects | `create_table`, `create_chart`, `create_summary_table` |
| Macros | `read_vba` |

Every parameter is documented in [TOOLS.md](TOOLS.md).

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

- Formulas are stored, not calculated. `read_range` in `values` mode returns the results
  Excel last saved, so formulas written by this server read as empty until the file is
  opened and saved in Excel or LibreOffice.
- Legacy `.xls` and `.csv` files are not supported.
- `create_summary_table` writes a static summary; openpyxl cannot create real PivotTables.
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
