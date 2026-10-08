# Changelog

All notable changes to this project are documented here. The format follows
[Keep a Changelog](https://keepachangelog.com/en/1.1.0/) and the project uses
[Semantic Versioning](https://semver.org/).

## [Unreleased]

### Security

- `import_workbook` checked only cell formulas, so a blocked function or a link to another
  workbook could be stored in a defined name (including a print area), a conditional
  format, a data validation rule or an Excel 2010 extension. It now checks every formula in
  the uploaded file. GHSA-69rc-6f6g-43xj (@NotAFlightRisk). Workbooks whose data
  validation uses `INDIRECT` (a common way to build dependent drop-down lists) or `CELL`
  are now rejected on import, as those rules already are when written through the tools.
- The formula check missed links to another workbook written without square brackets
  (`=Budget.xlsx!Sales`, `='\\server\share\book.xlsx'!Sales`), references to another
  workbook that end in a function call, and blocked functions after a range operator
  (`=SUM(A1:INDIRECT("B2"))`). A reference can now only name sheets of the workbook it is
  written to, so a formula that refers to a sheet that does not exist yet is rejected
  until that sheet is created. GHSA-frr9-2j4w-q923 (@zachary-satterly).

## [1.1.1] - 2026-09-28

### Fixed

- `read_vba` hides procedure attributes such as `Attribute Macro1.VB_ProcData...`, as the
  VBA editor does.

### Changed

- The README setup instructions cover current clients: Claude Desktop, Claude Code,
  Cursor, VS Code, OpenAI Codex, Gemini CLI and Devin Desktop.

## [1.1.0] - 2026-09-28

### Added

- `read_vba` shows the VBA macro code in `.xlsm` and `.xltm` workbooks, module by module,
  with each module's kind (standard, class, document or form). The code is only read,
  never run. `describe_workbook` reports `has_vba`.

### Changed

- The license names the author in full.

## [1.0.0] - 2026-09-28

A rewrite on the MCP Python SDK 2 and the MCP specification 2026-07-28, with a redesigned
tool set, a security hardening pass and a test suite.

### Security

- Formulas are checked on every write path, including `write_range`, conditional formats
  and data validation. The old check could be bypassed with lowercase names, a space before
  the parenthesis, or by writing through `write_data_to_excel`. The new check tokenizes
  formulas, normalizes function names, blocks a wider set of functions (including Excel 4
  macro functions and Google Sheets `IMPORT*` functions) and rejects DDE links and
  references to other workbooks. GHSA-vx23-x2cw-j2pq (@ACD421), GHSA-hcx2-gw6f-x9fq
  (@kkkh1), GHSA-h9g4-ghf5-r372 (@cawa102), GHSA-fqrm-pw9q-4hx4 (@arpitjain099),
  GHSA-wxp4-xc9x-vpxj (@hackchang), GHSA-9h7m-8h6c-6jh7 (@geo-chen), #119 (@joergmichno)
  and #134 (@EmersonZh).
- Every operation on a range is limited to 100,000 cells, and reads are paged, so a request
  for the whole grid no longer hangs the server. GHSA-27w4-5876-h4xq (@EQSTLab, @min8282)
  and GHSA-9r4q-rmpw-qv95 (@EQSTLab).
- Paths are confined to the `--allow-dir` folders after resolving symlinks, only Excel file
  extensions are accepted, and existing workbooks are never overwritten unless requested.
  GHSA-9h7m-8h6c-6jh7 (@geo-chen), #119 (@joergmichno), #115 (@starbuck100) and #128
  (@8endit).
- Streamable HTTP binds to `127.0.0.1` by default, enables DNS rebinding protection,
  supports a bearer token (`EXCEL_MCP_AUTH_TOKEN`) and refuses a public `--host` without
  one. #145 (@pzr21).
- Opened files are limited in size (`--max-file-mb`), and workbooks are saved atomically,
  so a failed save never corrupts the original file.

### Added

- `list_workbooks`, `export_workbook` and `import_workbook` for finding files and moving
  them to and from a remote server (#20, #62, #75).
- `find_cells`, `clear_range`, `set_sheet_layout` (column widths and autofit, row heights,
  frozen panes, auto filter, tab color; #67), `add_conditional_format` and
  `add_data_validation`.
- `read_range` modes for formula results or formula text (#94, #118), and paging for
  large ranges.
- Structured output with output schemas for every tool that returns data (#101).
- `--read-only` mode, `--max-file-mb` and `--log-level` options.
- Dates are read and written as ISO 8601 strings.
- Pictures in workbooks are kept when a workbook is edited (Pillow is now a dependency).
- MCP Registry metadata (`server.json`), a Dockerfile, and an MCPB bundle that runs with uv.
- CI on Linux, macOS and Windows for Python 3.11 to 3.14, with lint, type checks, tests
  and a dependency audit.

### Changed

- The tool set is redesigned. Old tools map to new ones as follows:

  | 0.x tool | 1.0 tool |
  | --- | --- |
  | `create_workbook` | `create_workbook` (refuses to overwrite unless `overwrite=true`) |
  | `get_workbook_metadata` | `describe_workbook` |
  | `create_worksheet` | `create_sheet` |
  | `copy_worksheet`, `rename_worksheet`, `delete_worksheet` | `copy_sheet`, `rename_sheet`, `delete_sheet` |
  | `read_data_from_excel` | `read_range` |
  | `write_data_to_excel`, `apply_formula` | `write_range` |
  | `validate_formula_syntax` | removed; `write_range` validates formulas |
  | `format_range` | `format_range` (options move into a `style` object) |
  | `merge_cells`, `unmerge_cells` | `merge_cells` with `action` |
  | `get_merged_cells`, `get_data_validation_info`, `validate_excel_range` | `describe_sheet` |
  | `copy_range` | `copy_range` |
  | `delete_range` | `clear_range`, or `delete_rows_or_columns` to shift cells |
  | `insert_rows`, `insert_columns` | `insert_rows_or_columns` |
  | `delete_sheet_rows`, `delete_sheet_columns` | `delete_rows_or_columns` |
  | `create_table` | `create_table` |
  | `create_chart` | `create_chart` (adds `column` and `bar` as separate types) |
  | `create_pivot_table` | `create_summary_table` |

- Failures are reported as MCP tool errors (`isError: true`) with actionable messages
  instead of "Error: ..." text (#144, #157).
- Every parameter is typed and documented, which fixes clients such as Gemini CLI and
  OpenCode that rejected untyped schemas (#89, #114, #158).
- `describe_workbook` reports the used range from cells that hold values, not cells that
  only carry formatting (#96).
- `create_summary_table` writes to the sheet and cell you choose, supports an aggregation
  per value column and no longer deletes sheets (#159).
- Charts show their axes in current Excel versions, including area charts already in a
  workbook, which openpyxl used to hide on every save.
- Concurrent edits to the same workbook are applied one after another instead of
  overwriting each other.
- Invalid files, including CSV and legacy `.xls` files renamed to `.xlsx`, give a clear
  error (#35).
- Configuration uses `--host`, `--port`, `--allow-dir` or the `EXCEL_MCP_*` variables.
  `FASTMCP_HOST` and `FASTMCP_PORT` are no longer read. `EXCEL_FILES_PATH` still works and
  now applies to stdio too.
- The server logs to stderr and no longer writes `excel-mcp.log`.
- Requires Python 3.11 or newer and the MCP Python SDK 2.

### Removed

- The SSE transport, deprecated by the MCP specification. Use Streamable HTTP.
- The unused `fastmcp` and `typer` dependencies.

### Thanks

Ideas and fixes from these pull requests were reimplemented in this release: #78
(@sawyer-shi), #90 (@davideconsonni), #93 (@ktan-bsp), #107 (@xanathar), #111
(@SOVALINUX), #113 (@RinZ27), #133 (@yagna-1), #136 (@JacobBrackett), #144 and #150
(@sanjibani), #157 (@Int-0X7FFFFFFF), #163 (@zk-hypersolid), and #165 and #166
(@vishalhabib99).

## [0.1.8] and earlier

See the [GitHub releases](https://github.com/haris-musa/excel-mcp-server/releases).

[Unreleased]: https://github.com/haris-musa/excel-mcp-server/compare/v1.1.1...HEAD
[1.1.1]: https://github.com/haris-musa/excel-mcp-server/compare/v1.1.0...v1.1.1
[1.1.0]: https://github.com/haris-musa/excel-mcp-server/compare/v1.0.0...v1.1.0
[1.0.0]: https://github.com/haris-musa/excel-mcp-server/compare/v0.1.8...v1.0.0
[0.1.8]: https://github.com/haris-musa/excel-mcp-server/releases/tag/v0.1.8
