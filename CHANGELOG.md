# Changelog

All notable changes to this project are documented here. The format follows
[Keep a Changelog](https://keepachangelog.com/en/1.1.0/) and the project uses
[Semantic Versioning](https://semver.org/).

## [Unreleased]

### Security

- `import_workbook` no longer stores uploads as macro-enabled workbooks (`.xlsm`, `.xltm`)
  unless the server is started with `--allow-vba-write`, the same rule `create_workbook`
  follows, whatever the upload holds (GHSA-px2f-qx58-w2qr, reported by
  @anelyberdecias-oss). Uploads to `.xlsx` and `.xltx` paths still have macros removed.
- `create_chart` now checks the size of the ranges it is given, and the number of series they
  make, before it builds anything, so one request cannot make the server work for minutes
  (GHSA-pwqg-v6rw-x2xh, reported by @anelyberdecias-oss). Chart ranges are limited by
  `--max-cells` like other ranges, and one call to `set_sheet_layout` can change at most 10,000
  rows in total.
- Slicer and timeline names are written to the workbook escaped, and names read back from a
  file are unescaped, so a name can no longer change the structure of the XML around it
  (GHSA-xw9m-c2v4-2hhq, reported by @anelyberdecias-oss). Slicer names and captions are limited
  to 255 characters as in Excel, and every tool refuses text with characters that XML cannot
  hold (control characters other than tab and line break).

## [2.2.4] - 2026-10-31

### Changed

- Saving a workbook after a cell patch is fast again on huge sheets that have only a few cells
  to patch (dynamic array cells): a few cells are looked up by address instead of scanning
  every cell, which GHSA-5px3-v5xm-4jhw's fix made about 1.3 s slower per 2 million cells.

### Fixed

- A workbook with a sparkline that has no location (`<xm:sqref>`) no longer fails with
  "Unexpected error" and a logged traceback when rows or columns are edited or sparklines are
  deleted. Excel itself cannot open such a file, so the server refuses it with a clear error.
  Ignored errors and protected ranges without `sqref`, sort conditions without `ref`, and a
  malformed list of slicer caches are refused the same way.

## [2.2.3] - 2026-10-30

### Security

- A crafted workbook can no longer make saving it take quadratic time (GHSA-5px3-v5xm-4jhw,
  reported by @q7256). Conditional formats of Excel 2010 that the server cannot restore are
  removed in one pass over the sheet's extensions instead of one pass per rule. The same flaw
  is fixed wherever preserved content was rescanned per item: elements that close themselves
  in sparkline, conditional format and validation extensions, the relationships and part names
  of a package, cells with dynamic array metadata, and slicer caches that no slicer uses.

## [2.2.2] - 2026-10-21

### Security

- A small workbook can no longer make the server use gigabytes of memory by holding millions of
  cells: the size limits counted bytes, and memory depends on cells (GHSA-5p2h-j6m4-5fm2,
  reported by @arpitjain099). The cells and rows of each worksheet, and the cells of merged and
  hyperlink ranges, are now counted before the sheet is parsed, and a workbook that would create
  more than 3,000,000 is refused. `--max-loaded-cells` (`EXCEL_MCP_MAX_LOADED_CELLS`) sets the
  limit. Reads that stream (`read_range`, `export_csv`) are not affected.
- Pillow, used to read images for `insert_image`, is required at 13.0 or later, which fixes
  memory-safety and denial-of-service issues in its image decoders (CVE-2026-112895 and
  others); 2.2.1 allowed only 12.x.

### Changed

- Sheets whose rows or cells carry no `r` attribute, as some writers make them, are edited
  through the large-sheet path like any other, instead of being loaded whole.
- `describe_sheet` of a large sheet no longer parses every cell: it reads the notes, links, hidden
  rows and used range from the rows that carry them and from the sheet's XML, with the same result.
  A `create_pivot_table` source whose formulas have no stored results, which needs the cells in
  memory to be calculated, says to open and save the workbook in Excel when it is over the limit.

## [2.2.1] - 2026-10-18

### Changed

- Edits that look at every cell of a large sheet no longer load it whole. On a 200,000-row
  sheet with four filled-down formula columns (Excel's file / one written by openpyxl):
  inserting or deleting rows took 140-170 s and now takes 9 / 16 s, inserting or deleting columns
  93-218 s and now 10-13 / 13-18 s, `replace_cells` about 100 s (any text) and now 4 s when the text is
  rare (40 / 32 s when it changes 200,000 formulas), `sort_range` of 100,000 rows 31 / 25 s and
  now 21 / 15 s, `copy_sheet` of 950,000 cells 94 s and now 4 s, and `create_pivot_table`
  108 s (50,000 rows: 28 s) and now 19 s (8 s). Their results are the ones loading
  the whole sheet gave. How:
  - Rows that only move are renumbered in their XML, and the references in their formulas are
    rewritten as text; a group of filled-down formulas that the edit cuts, or whose references
    it cuts, is parsed and becomes ordinary formulas, as Excel splits it. Only the rows around
    the edit are parsed.
  - `replace_cells` parses only the blocks of rows that may hold the text. A copy of a sheet
    shares the original's rows that were not parsed. A PivotTable source is read as the file
    stores it (the stored results of formulas) when every formula has one.
  - Rows that are parsed are read and written as text instead of through an XML library, 2 to 3
    times faster, in the forms Excel and openpyxl write; any other form goes through openpyxl.
  - Formulas made only of plain cell references are moved, translated and checked by the formula
    check once per shape instead of by the tokenizer one by one. The check gives the same
    answers (tests compare it with the tokenizer on random formulas).
- Inserting columns creates the formats of the inserted cells once per kind of neighbour
  instead of once per cell, which was most of the 130 s that inserting two columns into a
  200,000-row sheet took.
- `copy_sheet` of a sheet over the cell limit is refused without loading the sheet (it took
  31 s to count its cells). `rename_sheet` of a sheet that a large sheet refers to, and
  `create_pivot_table` on the sheet of its source, no longer parse every row of the large
  sheet.

## [2.2.0] - 2026-10-18

### Added

- `import_csv` reads a `.csv` or `.tsv` file into a sheet (new or existing) at a cell, and
  `export_csv` writes a range as a CSV file. Import detects the delimiter (or takes `comma`,
  `semicolon`, `tab`, `pipe`), reads UTF-8 with or without a BOM, UTF-16 and `cp1252`, keeps
  quoted fields and line breaks in them, and types fields as Excel's import does: numbers
  (also `1,234.50` and `12%`; `decimal` for a decimal comma), `TRUE`/`FALSE` and ISO dates
  become values. Text Excel would alter stays text: leading zeros, more than 15 digits and
  numeric dates such as `01/02/2026`, which are read only in the `date_order` you name.
  `as_text` keeps everything as text. Export writes raw values (ISO dates) or, with
  `formatted`, what Excel shows, with any delimiter and encoding including `utf-8-sig` (a BOM
  for Excel). Both stay within `--max-cells` and the file size limit, and go through the path
  gate (only `.csv` and `.tsv` files).
- CSV formula injection: imported fields are never formulas. Text starting with `=`, `+`, `-`
  or `@` (or a tab or carriage return) is stored as text with Excel's quote prefix, so editing
  the cell does not run it; a field that is a number, such as `-5`, is a number. On export,
  text cells starting with those characters get a `'` prefix (`escape_formulas`, on by
  default); numbers never do.
- `--tools AREAS` (`EXCEL_MCP_TOOLS`) offers only the core tools (`describe_workbook`,
  `describe_sheet`, `read_range`) and the chosen areas, to cut the context the tool list uses.
  Areas: `workbooks`, `csv`, `sheets`, `data`, `formulas`, `formatting`, `tables`, `charts`, `pivots`,
  `slicers`, `names`, `notes`, `images`, `vba`. It combines with `--read-only` and
  `--allow-vba-write` (VBA writing needs the flag and the `vba` area). An unknown area stops the
  server at startup with the valid names.
- `trace_formula` audits formulas like Excel's Formulas tab, calculated by the built-in engine
  (read-only): `precedents` (the cells and ranges a formula uses) and `dependents` (the
  formulas that use a cell or range), one level or all, across sheets, defined names, table
  references and spill references, merged into A1 ranges by sheet, with `INDIRECT` and
  `OFFSET` followed to the cells they compute and the formulas where that cannot be worked
  out listed as `unknown`; and `evaluate`, the steps of Evaluate Formula (each
  sub-expression with its value, innermost first, only the branches Excel takes). Results are
  capped at 200 ranges and 100 steps, with `more` counting what was left out.
- `goal_seek` is Excel's Goal Seek: it changes a constant until a formula reaches a value
  (within 0.001, at most 100 iterations) and stores the answer, or changes nothing and says
  that it may not have found a solution. It refuses a target without a formula, an input that
  holds a formula or text, a formula that does not depend on the input, and a formula the
  calculator cannot calculate. Results are checked against `Range.GoalSeek` and
  `Range.DirectPrecedents`, `Precedents`, `DirectDependents` and `Dependents` recorded from
  Excel (`tests/fixtures/audit_golden.json`).

### Changed

- Edits and descriptions of large workbooks are much faster. A sheet of 1 MB or more (of XML)
  is loaded without its rows: an edit parses only the blocks of 128 rows it touches, and the
  save writes those and copies the other rows as the file has them (their stored formula
  results are dropped, as for any edit). A 10-cell write to a 200,000-row sheet went from
  about 50 s to 3 s (50,000 rows: 11 s to under 1 s), and `describe_workbook` finds the used range in the sheet's XML text instead of
  parsing every cell (200,000 rows: 12 s to 1.4 s). Sheets with array or data-table formulas or cell
  metadata, and macro-enabled workbooks, are still loaded whole; so are edits that look at
  every cell (inserting rows, sorting, replacing).
- Sheets with shared formulas, which Excel writes whenever a formula is filled down, get the
  same fast edits, so most large workbooks made in Excel do. A block of rows is translated from
  the master formula wherever it is, and a master in a block that was parsed is written as a
  master again, so the rows kept from the file still refer to it. Overwriting a master parses
  the blocks of its range, and its formulas become ordinary ones, as when every row is parsed.

### Fixed

- `=EXP(1000)` and other overflowing exponentials are `#NUM!`, as in Excel, instead of
  leaving the formula uncalculated.

## [2.1.1] - 2026-10-16

### Fixed

- `create_pivot_table` (and the slicers and timelines that make PivotTables again) accept the
  sources Excel pivots without complaint, and show what Excel shows: columns that mix text,
  numbers, dates, booleans and blanks; cells with errors or formulas (their stored results, or
  the calculator's); times and durations. Row, column and filter fields list every kind of
  item in Excel's order; sum, average, min and max use the numbers (and dates, as serial
  numbers) and skip the rest, count counts what is not blank, and an error among the values
  is the result, including in the counts of totals. The cache holds the flags and items Excel
  writes (`containsMixedTypes`, `b`, `e`, `d` ...), a field's number format is the one all its
  cells share, and a timeline needs a field of dates only. Grouping a mixed field is refused
  as in Excel ("Cannot group that selection"). Fixes #215.

### Changed

- Editing a workbook that holds a large PivotTable no longer parses its cache records: any tool
  on a 134,000 by 13 source now takes as long as on the same workbook without the PivotTable
  (`add_slicer` 93 s to 40 s, `write_range` 90 s to 36 s, `describe_sheet` 60 s to 19 s) and
  peaks at 1 GB instead of 2.9 GB. The records are kept as they are in the file; a slicer on a
  field without shared items reads that one column from them.

## [2.1.0] - 2026-10-16

### Added

- `--max-pivot-cells` (`EXCEL_MCP_MAX_PIVOT_CELLS`) sets the largest PivotTable source, and
  `--max-cells` (`EXCEL_MCP_MAX_CELLS`) the most cells one call may read, write or format.
  `--max-file-mb` now has `EXCEL_MCP_MAX_FILE_MB`. The error for an oversized source names the
  option that raises it (#210).

### Changed

- PivotTable sources may have up to 3,000,000 cells (was 100,000, the general range limit), so
  a sheet of 134,000 rows by 13 columns can be summarized. Creating a PivotTable from such a
  source is faster (92 s to 54 s) and uses half the memory (2.1 GB to 1.0 GB), because the
  pivot cache records are written as XML text instead of one object per value.

### Fixed

- TOOLS.md now documents every nested parameter (arrays of objects, optional objects and
  objects within objects, such as `create_pivot_table`'s `value_fields`) with each field's
  type, requirement, default, allowed values and description, and describes `path`, `sheet`,
  `range`, `at` and `cell` (#209).

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

[Unreleased]: https://github.com/haris-musa/excel-mcp-server/compare/v2.2.4...HEAD
[2.2.4]: https://github.com/haris-musa/excel-mcp-server/compare/v2.2.3...v2.2.4
[2.2.3]: https://github.com/haris-musa/excel-mcp-server/compare/v2.2.2...v2.2.3
[2.2.2]: https://github.com/haris-musa/excel-mcp-server/compare/v2.2.1...v2.2.2
[2.2.1]: https://github.com/haris-musa/excel-mcp-server/compare/v2.2.0...v2.2.1
[2.2.0]: https://github.com/haris-musa/excel-mcp-server/compare/v2.1.1...v2.2.0
[2.1.1]: https://github.com/haris-musa/excel-mcp-server/compare/v2.1.0...v2.1.1
[2.1.0]: https://github.com/haris-musa/excel-mcp-server/compare/v2.0.0...v2.1.0
[2.0.0]: https://github.com/haris-musa/excel-mcp-server/compare/v1.1.2...v2.0.0
[1.1.2]: https://github.com/haris-musa/excel-mcp-server/compare/v1.1.1...v1.1.2
[1.1.1]: https://github.com/haris-musa/excel-mcp-server/compare/v1.1.0...v1.1.1
[1.1.0]: https://github.com/haris-musa/excel-mcp-server/compare/v1.0.0...v1.1.0
[1.0.0]: https://github.com/haris-musa/excel-mcp-server/compare/v0.1.8...v1.0.0
[0.1.8]: https://github.com/haris-musa/excel-mcp-server/releases/tag/v0.1.8
