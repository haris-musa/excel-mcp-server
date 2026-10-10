# Changelog

All notable changes to this project are documented here. The format follows
[Keep a Changelog](https://keepachangelog.com/en/1.1.0/) and the project uses
[Semantic Versioning](https://semver.org/).

## [Unreleased]

## [2.0.0] - 2026-10-11

2.0 grows the server from 26 to 42 tools: charts of every Excel type, PivotTables, slicers,
sparklines, sorting, find and replace, print setup, protection and opt-in VBA writing. It
calculates formulas without Excel, updates references on every structural edit as Excel does,
and keeps content it cannot model instead of deleting it. Tool parameters and results are now
consistent, which breaks 1.x clients; see "Upgrading from 1.x" at the end.

### Added

- Tools: `set_table`, `create_pivot_table`, `delete_pivot_table`, `delete_chart`,
  `add_sparklines`, `delete_sparklines`, `add_slicer`, `delete_slicer`, `sort_range`,
  `transform_range`, `replace_cells`, `set_defined_name`, `delete_defined_name`, `set_note`,
  `delete_note`, `insert_image`, `delete_image` and `set_workbook_settings`.
- Opt-in VBA writing: `--allow-vba-write` (or `EXCEL_MCP_ALLOW_VBA_WRITE=1`) adds
  `write_vba_module` and `delete_vba_module` for `.xlsm`/`.xltm` files (standard, class,
  `ThisWorkbook` and sheet modules; creates the VBA project if needed). Absent by default and
  in `--read-only` mode; code is stored, never run. `create_workbook` can then create
  `.xlsm`/`.xltm`.
- Formula calculation: `read_range` in `values` mode calculates formulas with no stored result
  (about 260 functions, including `LET`, `SORT`, `FILTER`, `UNIQUE`, `SEQUENCE`, `SUBTOTAL`,
  `AGGREGATE`, `INDIRECT`, `CELL`, `INFO`, `HYPERLINK`, structured references and spill
  references `A1#`), checked against 2,800+ results recorded from Excel. Results Excel stored
  are kept; formulas it cannot reproduce are returned as null and listed in `uncalculated`.
  Work per call is bounded. `read_range` also accepts whole-column and whole-row spans
  (`B:B`, `2:3`).
- Formulas are checked against Excel's syntax and argument counts before anything is written,
  in every place a formula can be stored (`Formula '=SUM(A1:' is not valid: unclosed '('.`).
  Functions added after 2007 are stored with the `_xlfn.` prefix. Formulas that return several
  values are stored as dynamic array formulas, as Excel stores them, and every save
  recalculates their spill ranges (`#SPILL!` where cells are in the way); `write_range`
  reports formulas that a non-empty cell blocks in `blocked`.
- `create_chart`: any ranges on any sheet through `series` (or a `source` block, with
  `series_in`), `replace` and `name`, a chart sheet when `at` is omitted; types `doughnut`,
  `radar`, `bubble` and the Excel 2016 `waterfall`, `histogram`, `pareto`, `box_whisker`,
  `treemap`, `sunburst` and `funnel`; scatter subtypes (`scatter_style`); combo charts and
  secondary axes; trendlines and error bars; axis title, scale, `log`, `reverse`, number
  format, gridlines and label position; data labels; series colors and markers; `grouping`,
  `colors`, `legend`, `title_size`, `plot_color` and `style`. Options that do not fit the chart
  type are rejected.
- `create_pivot_table` makes a real PivotTable that Excel refreshes, with its figures also
  written to the cells (matching Excel's, checked against about 200 recorded PivotTables):
  `show_as` and `number_format` per values field, `field_settings` (`show_items`, `sort`,
  `group_dates`, `group_numbers`), `calculated_fields`, `layout`, `subtotals`, `values_in`.
- `add_slicer` and `delete_slicer`: slicers for tables and PivotTables and timelines for
  PivotTable date fields, with selection (filtering the PivotTable or table), style, columns,
  sort, and one slicer for several PivotTables that share a cache. Grouped dates work.
- `add_sparklines`: line, column and win/loss sparklines with colors, emphasized points, axis
  options and date axes.
- `set_table` takes `options` like Excel's Table Design tab (style, header row, totals row with
  a function per column, banded rows and columns, first and last column, filter buttons),
  calculated columns (`=[@Price]*[@Qty]`, filled into new rows) and resizing through `range`.
- `add_conditional_format`: the rule types of Excel's menu (`top`, `bottom`, `above_average`,
  `below_average`, `duplicate`, `unique`, `contains_text`, `not_contains_text`, `begins_with`,
  `ends_with`, `date`, `blanks`, `no_blanks`, `errors`, `no_errors`, `icon_set` with the 17
  standard sets plus `3Stars`, `3Triangles` and `5Boxes`, custom `icons`, `thresholds`,
  `reverse`), Excel 2010 data bars (gradient or solid, border, negative bars, axis),
  `hide_values`, `stop_if_true` and `priority`.
- `add_data_validation`: list `source` from cells or a name, `time` type, input message
  (`prompt_title`, `prompt`), `error_style`, `error_title`, and dates written as `2026-01-31`.
- `set_sheet_layout`: hide, show and group rows and columns (`rows`, `columns`); sheet
  `visibility`, `position` and `view` (zoom, gridlines, headings, right to left, active sheet,
  selected cell); `print_setup`; sheet `protection`; and `auto_filter` with criteria (values,
  comparisons, top or bottom N, average, fill color) that hides the rows that fail. A table is
  filtered by passing its name as `range`.
- `write_range` takes `links`: hyperlinks to `http`, `https`, `mailto` or places in the
  workbook, with a tooltip. `format_range` sets `locked` and `formula_hidden`. `clear_range`
  has `clear: "rules"` (conditional formats and validation).
- `copy_range` takes `paste` (`all`, `values`, `formulas`, `formats`), `transpose` and
  `skip_blanks`. `transform_range` runs `remove_duplicates`, `text_to_columns` and `fill`
  (Fill Down, Right and Series).
- `describe_sheet` lists charts, images, PivotTables, slicers, timelines, sparklines, notes,
  hyperlinks, hidden rows and columns, the print area, protection and view, each object by
  `name`. `describe_workbook` also returns document properties, calculation settings, the
  active sheet, `chart_sheets` and object counts per sheet.
- Editing keeps content that openpyxl cannot model, byte for byte: sparklines, Excel 2010
  conditional formats and validation, slicers and timelines, newer chart types, threaded
  comments, dynamic array metadata, shapes, form controls, header and footer pictures, custom
  XML and PivotTable extensions. Content whose target is deleted follows Excel (an unused
  slicer cache is removed, validation formulas become `#REF!`).
- `Limits` gains `max_copy_cells` (`copy_sheet`), `max_unpack_factor` and
  `max_compression_ratio`.

### Changed

- Built on the MCP Python SDK 2.3, which supports the 2026-07-28 protocol revision. Tool
  schemas write unions as `anyOf` of single types instead of `"type": [...]` lists, which some
  clients (Gemini, strict validators) reject.
- **Breaking:** consistent parameters, one name per concept: `sheet`, `range`, `at` (the
  top-left cell of whatever is placed), `source`, `name`, with units in names (table below).
  Charts and images are selected by `name`, not by index.
- **Breaking:** `set_table` replaces `create_table` and also changes an existing table
  (`style` and `striped_rows` moved into `options`; new tables use `TableStyleMedium2`).
  `create_pivot_table` replaces `create_summary_table`; its result is a PivotTable.
- **Breaking:** `create_chart` loses `show_legend`, `data_sheet` and the options
  `x_axis_title` and `y_axis_title`; `options.legend` is `bottom`, `right`, `left`, `top` or
  `none` and defaults to Excel's for the chart type; `at` is optional. `create_chart` is
  marked destructive, since `replace` overwrites a chart.
- **Breaking:** charts are drawn as current Excel draws them: gray text, light gridlines, no
  rounded corners, one color per series, gaps between clustered columns. Scatter charts plot
  points, line charts are straight without markers unless `smooth` or `markers` is set, and
  horizontal bar charts list rows top-down.
- **Breaking:** compact results. `read_range` returns `{range, values, next_range?}` (no
  `sheet` or `truncated`; page on while `next_range` is present) and omits trailing empty
  cells and rows. `find_cells` returns `matches` as `{sheet: {cell: value}}`.
  `describe_workbook` returns `sheets`, `defined_names` (objects with `name`, `refers_to`,
  `sheet`; hidden names left out) and `has_vba`, without path, size or counts.
  `list_workbooks` returns `{path: size_bytes}`. `describe_sheet` drops `chart_count` (use
  `charts`) and empty fields. Midnight dates read as `2026-01-31`.
- **Breaking:** tools that change a workbook return a small object (`sheet`, `range`, `name`,
  and a `note` where needed) instead of a sentence, and list no output schema; only tools
  that return data keep one.
- **Breaking:** `set_sheet_layout` takes `column_widths_chars` and `row_heights_pt` as objects
  (`{"A": 20}`, `{"1": 30}`), and `auto_filter` is an object (`range`, `filters`, `remove`).
- **Breaking:** `add_conditional_format` rules that format cells need a `fill_color` or
  `font_color`, and fields that do not fit the rule type are rejected. New rules take the next
  free priority.
- **Breaking:** `add_data_validation` `error_message` and `prompt` are limited to 255
  characters, and a list needs `options` or `source`.
- `insert_rows_or_columns`, `delete_rows_or_columns` and `rename_sheet` update references as
  Excel does, workbook-wide: formulas, names, conditional formats, validation, merges, tables,
  filters, print settings, charts, images, PivotTables, sparklines and shapes. Like Excel,
  an edit that would cut through an array formula, a PivotTable or a table header fails and
  saves nothing. Rows inserted into a table get its calculated column formulas.
- `copy_sheet` copies what Excel's "Create a copy" does: validation, conditional formats,
  images, charts (re-pointed at the copy), tables (renamed), PivotTables, slicers and
  timelines, shapes, form controls, embedded objects, protected ranges, freeze panes, filters,
  print setup, protection and sheet-scoped names. What it cannot copy is named in the
  result's `note`.
- `INDIRECT`, `CELL` and `INFO` are allowed in formulas, since they only read the open
  workbook and the host. `HYPERLINK` is allowed with a literal `http://`, `https://`,
  `mailto:` or `#Sheet!A1` link.
- `read_range`, `find_cells` and `describe_workbook` stream the workbook: memory stays flat
  instead of growing with the sheet, and formulas without stored results are calculated from
  the rows they use. `import_workbook` checks uploads without loading every cell.
- Invalid arguments are reported as one readable line per problem, with suggestions
  (`rnage: unknown field; did you mean 'range'?`). Invalid formula syntax fails with
  `InvalidFormulaError`.
- An edit whose saved file would exceed the file size limit is refused and leaves the file
  unchanged.
- Tool schemas are much smaller: no generated titles, `null` unions or defaults for absent
  values, and `path`, `sheet` and range parameters are explained once in the server
  instructions. A test keeps the total under a budget.

### Removed

- **Breaking:** `create_table` (use `set_table`) and `create_summary_table` (use
  `create_pivot_table`).

### Fixed

- `write_range` stored formulas that return several values (`=SUM(A1:A5*2)`) as plain
  formulas, which Excel showed with `@` and a single value; they are now dynamic array
  formulas.
- Formulas using `IFS`, `XLOOKUP`, `TEXTJOIN`, `STDEV.S`, `SORT` and other newer functions
  showed `#NAME?` in Excel (`SORT` made the file unopenable) because the `_xlfn.` prefix was
  missing. Formulas are shown as typed, without `_xlfn.`, `_xlws.` or `_xlpm.`.
- Editing a workbook silently deleted sparklines, slicers, timelines, threaded comments,
  newer chart types, shapes, form controls and the other content listed under Added, and
  damaged charts (chart style, rounded corners, plot area fill, area chart axes, empty labels
  shown as "None"). With several PivotTable caches, every cache after the first was linked to
  the first one's records.
- Chart titles, axis titles and legends sat on top of the plot in Excel.
- Date limits in data validation (`2026-01-31`) were stored as the formula `2026-01-31`.
- Inserting or deleting rows or columns left hyperlinks in place; they move with their cells,
  and `copy_range` copies them.
- `set_sheet_layout` could write overlapping column definitions.
- `format_range` `font_name` took no effect in Excel on files this server creates; setting a
  name now clears the theme font scheme, as picking a font in Excel does.
- Dates written with `write_range` into a default-width column showed `####`; the column now
  widens to fit.
- `import_workbook` into an `.xlsx` or `.xltx` path left an orphan VBA project; it is removed,
  as Excel does.
- New workbooks named "openpyxl" as their creator.
- A drive-relative path such as `C:book.xlsx` gets a clear error.
- A tool that fails unexpectedly reports "Unexpected error in <tool>; see server log" and logs
  the exception, instead of the SDK's "Error executing tool <tool>".

### Security

- Every workbook the server opens is checked against the zip bomb limits
  (`max_unpack_factor`, `max_compression_ratio`), not only uploads. GHSA-jf2g-fjwq-rjr6
- `import_workbook` checks every XML part wherever it is stored; rejects duplicate or unsafe
  part names, data connections, query tables, external links, linked OLE objects and external
  relationships other than hyperlinks; caps expanded size and compression ratio while
  reading; and scans in constant memory with a nesting limit. Hyperlinks to files, network
  locations, schemes other than `http`, `https` and `mailto`, or with credentials are removed
  and reported in `note`. GHSA-99mr-c7r6-xhrc
- Network (UNC) paths, Windows device paths and reserved device names are rejected, with or
  without `--allow-dir`, unless the allowed folder is on that share; the rules apply again to
  the resolved path. GHSA-wgvf-cmgr-7573
- Excel 4.0 macro functions that read files, the system or workbook internals, or run other
  programs (`FILES`, `GET.*`, `APP.*`, `RUN`, `EXEC`, `SEND.KEYS`, `SQL.*`, `MAIL.*` and
  related) are blocked in every formula, including defined names. GHSA-w7wm-p57p-8q22
- The formula calculator matches wildcards without backtracking, limits formula length and
  nesting, checks sizes before building text and arrays, and ends hostile formulas as
  uncalculated cells.
- Reading a VBA project fails fast on malformed containers and caps the module count and the
  decompressed size. GHSA-mvpx-mwxr-g2f5

### Upgrading from 1.x

| 1.x | 2.0 |
| --- | --- |
| `create_table` (`style`, `striped_rows`) | `set_table` (`options.style`, `options.striped_rows`) |
| `create_summary_table` (`source_range`, `group_by`, `values`, `target_sheet`, `target_cell`) | `create_pivot_table` (`source`, `row_fields`, `value_fields`, `sheet`, `at`); a value's `aggregation` is `function` |
| `create_sheet` `sheet` | `new_name` |
| `insert_rows_or_columns`, `delete_rows_or_columns` `at` | `start` |
| `write_range` `start_cell` | `at` |
| `copy_range` `target_cell`, `target_sheet` | `at`, `to_sheet` |
| `create_chart` `anchor_cell`, `data_range`, `data_sheet` | `at`, `source` (`Data!A1:C13`) |
| `create_chart` `options.show_legend`, `x_axis_title`, `y_axis_title` | `options.legend`, `options.x_axis.title`, `options.y_axis.title` |
| `set_sheet_layout` `column_widths`, `row_heights` (lists) | `column_widths_chars`, `row_heights_pt` (objects) |
| `set_sheet_layout` `auto_filter: "A1:F100"` | `auto_filter: {"range": "A1:F100"}` |
| `describe_sheet` `chart_count`, `column_widths` | `charts`, `column_widths_chars` |
| `describe_workbook` `defined_names` (strings) | objects with `name`, `refers_to`, `sheet` |
| `read_range` `sheet`, `truncated` | removed; page while `next_range` is present |
| `find_cells` `matches` as a list | `{sheet: {cell: value}}` |
| Results as sentences | objects with `sheet`, `range`, `name` |

## [1.1.2] - 2026-10-08

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
- Warnings from libraries such as openpyxl, which can quote workbook content (for example
  a defined name used as a print area), are no longer written to stderr. They are logged
  only with `--log-level DEBUG`.

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

[Unreleased]: https://github.com/haris-musa/excel-mcp-server/compare/v2.0.0...HEAD
[2.0.0]: https://github.com/haris-musa/excel-mcp-server/compare/v1.1.2...v2.0.0
[1.1.2]: https://github.com/haris-musa/excel-mcp-server/compare/v1.1.1...v1.1.2
[1.1.1]: https://github.com/haris-musa/excel-mcp-server/compare/v1.1.0...v1.1.1
[1.1.0]: https://github.com/haris-musa/excel-mcp-server/compare/v1.0.0...v1.1.0
[1.0.0]: https://github.com/haris-musa/excel-mcp-server/compare/v0.1.8...v1.0.0
[0.1.8]: https://github.com/haris-musa/excel-mcp-server/releases/tag/v0.1.8
