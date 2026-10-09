# Changelog

All notable changes to this project are documented here. The format follows
[Keep a Changelog](https://keepachangelog.com/en/1.1.0/) and the project uses
[Semantic Versioning](https://semver.org/).

## [Unreleased]

### Added

- `set_table` takes `options` like Excel's Table Design tab: a totals
  row with a function per column (sum, average, count, count numbers, max, min, standard
  deviation, variance, or a custom formula) or a label, written as Excel writes them
  (`totalsRowFunction`, `SUBTOTAL(109,...)` cells, `totalsRowLabel`); calculated columns from
  a formula with structured references such as `=[@Price]*[@Qty]` or relative cells such as
  `=B2*C2`, filled into every row, into rows added by `set_table` and into rows inserted
  later (`calculatedColumnFormula`); a new totals row is what Excel makes ("Total" first, and
  the last column summed or counted); resizing moves the totals row; header
  row on or off, banded columns, first and last column emphasis, filter buttons; and resizing
  (`range`). Checked against the same steps in Excel.
- The formula calculator reads structured references: `Sales[Price]`, `[@Price]`,
  `Sales[[#This Row],[Price]]`, `Sales[#Headers]`, `[#Data]`, `[#Totals]`, `[#All]`,
  combinations and column ranges. 34 cases recorded from Excel were added to its test fixture.
- `clear_range` has `clear: "rules"`, which removes conditional formats and data validation
  (Clear Rules and Data Validation > Clear All) from the range, splitting rules that cover more
  cells; `formats` also clears conditional formats and `all` also clears validation, as in
  Excel. Rules in Excel 2010 extensions (data bars, icon sets, cross-sheet validation) follow.

- `create_chart` makes the Excel 2016 chart types `waterfall` (`totals`, `connector_lines`),
  `histogram` (`bins`: width, count, underflow, overflow), `pareto`, `box_whisker` (`box`:
  quartile method, mean marker and line, inner points, outliers), `treemap`
  (`parent_labels`), `sunburst` and `funnel`, with title, legend, data labels and, where
  Excel has them, axis titles and scale. They are written as Excel writes them (chartex
  parts, hidden `_xlchart` names, style and colour parts) and are listed by `describe_sheet`
  after the other charts, removed by `delete_chart`, replaced with `replace` and copied by
  `copy_sheet`. Checked against the same charts made in Excel; Excel opens them without repair. A new
  chart looks like an inserted one: labels on waterfall and funnel bars (values) and on treemap
  and sunburst tiles (categories), overlapping treemap group labels; `data_labels` with
  `show: []` hides them. Data blocks are read as Excel reads them (leading text columns are
  labels, a single column of numbers is enough for box and whisker, histogram, funnel,
  waterfall, treemap and sunburst, and a Pareto of numbers is binned).
- `add_sparklines` and `delete_sparklines`: Insert > Sparklines with line, column and
  win/loss types, colors, emphasized points (markers, high, low, first, last, negative), axis
  and scale options, date axes, empty-cell handling and line weight. Sparklines in the target
  cells are replaced. `describe_sheet` lists them. They are written as Excel writes them, are
  kept by every edit and by `copy_sheet`, and follow rows and columns through the package
  references API.
- `add_conditional_format` data bars are written as Excel 2010 and later write them: gradient
  or solid fill, border, negative bar fill and border, axis position and color, bar direction,
  and automatic, lowest/highest, number, percent, percentile or formula ends. Icon sets accept
  the sets that exist in the extension only (`3Stars`, `3Triangles`, `5Boxes`) and custom
  `icons` per range, including no icon. `describe_sheet` lists them; `copy_sheet` copies them.
- `add_slicer` and `delete_slicer`: slicers for tables and PivotTables and timelines for
  date fields of PivotTables, written as Excel writes them (checked part by part against
  Excel's own files, and opened in Excel without repair). Options: field, caption,
  position and size, columns, built-in style, sort, hide items with no data, selected items
  or period, and one slicer connected to several PivotTables that share a cache. Selecting
  items filters like the buttons do: PivotTable items are hidden and the figures recalculated
  from the source (equal to Excel's after a refresh), table rows are filtered and hidden.
  `describe_sheet` lists slicers and timelines with their selection. On PivotTables made or
  refreshed by Excel the items are hidden in the definition and the cache is marked to refresh
  on load, so Excel recalculates when it opens the file (the tool result says so); checked
  against Excel's own slicer selections on Excel-made PivotTables.
- `copy_sheet` copies slicers and timelines as Excel does (a copy of a sheet with its
  PivotTables and slicers filters itself; slicers elsewhere gain the copied PivotTables).
- Deleting a table, or the column of a table slicer, with `delete_rows_or_columns` removes
  the slicer, as Excel does. `delete_slicer` keeps what Excel keeps (a table's filter, a
  row, column or filter field's hidden items) and clears a timeline's period and the hidden
  items of a field the PivotTable does not show.
- `read_range` in `values` mode calculates formulas that have no stored result, which used
  to read as null until Excel saved the file. A built-in calculator covers about 260
  functions (math, statistics, financial and securities, dates, text, lookup, logical,
  `LET`, `SORT`/`FILTER`/`UNIQUE`/`SEQUENCE`, `SUBTOTAL`, `AGGREGATE`). It follows Excel's
  rules for blanks, text-numbers, errors, criteria wildcards and rounding; results Excel
  stored are kept. Formulas it cannot reproduce exactly are returned as null and listed in
  `uncalculated` with the reason. It is checked against 2,500+ results recorded from real
  Excel, and its work per call is bounded.
- Formulas written with functions Excel added after 2007 are stored with the `_xlfn.`
  prefix (and `_xlpm.` for `LET` names): in cells, `copy_range`, `sort_range`, conditional
  formats, data validation and defined names.
- Formula chains of any length (running balances, amortization schedules) are calculated.
- `read_range` accepts whole-column and whole-row spans (`B:B`, `2:3`), limited to the used
  range.
- Formulas Excel cannot parse are rejected before anything is written, in every place a
  formula can be stored: unclosed or unmatched parentheses, a missing operand, an operator
  where a value belongs, unterminated text, malformed arrays, references and names that
  cannot exist, such as `Formula '=SUM(A1:' is not valid: unclosed '('.`. Checked against
  Excel: every accepted formula opens without repair, and the ones Excel refuses are
  refused too. Functions Excel knows must get an argument count it accepts, from a table
  generated by asking Excel (`scripts/probe_function_arguments.py`, Windows and Excel only).
- Spill references (`A1#`, `Sheet1!A1#`, `name#`) are accepted, stored as
  `_xlfn.ANCHORARRAY(A1)` as Excel's file format does, and shown as `A1#` by `read_range` in
  `formulas` mode.
- The calculator evaluates spill references: `=SUM(A1#)`, `=COUNTA(A1#)` and `=INDEX(A1#,2)`
  use the range the dynamic array formula in `A1` filled, including a spill of a spill, and a
  reference to a cell that holds no formula is `#REF!`, as in Excel.

- `add_data_validation`: dropdown lists can come from cells or a name (`source`, e.g.
  `=$A$2:$A$20`, `=Sheet2!$A:$A`, `=Regions`), the `time` type is supported, and rules take an
  input message title and text (`prompt_title`, `prompt`), an alert style (`error_style`:
  `stop`, `warning` or `information`), an alert title and message. Date and time limits can be
  written as `2026-01-31` and `09:30`; titles and messages are limited to Excel's 32 and 255
  characters. All of it passes the formula check.
- `add_conditional_format` offers the rule types of Excel's menu: `top` and `bottom` (N items
  or percent), `above_average` and `below_average` (equal, standard deviations),
  `duplicate`, `unique`, `contains_text`, `not_contains_text`, `begins_with`, `ends_with`,
  `date` (yesterday, last 7 days, this week, next month and others), `blanks`, `no_blanks`,
  `errors`, `no_errors` and `icon_set` (the 17 standard 3, 4 and 5 icon sets, with
  thresholds as percent, number or percentile, `reverse` and `icon_only`), plus
  `stop_if_true` and `priority`. Rules write the formulas Excel writes. Fields that do not
  fit the rule type are rejected.
- `set_sheet_layout`'s `auto_filter` applies criteria: values to show, comparisons (equals,
  greater than, begins with, contains and their negations, one or two joined by and/or),
  top or bottom N items or percent, above or below average, and fill color. Rows that fail
  are hidden, as Excel does, so the file opens filtered. Tables can be filtered by passing
  the table's name as `range`, and `remove` removes a filter and shows its rows.
- `copy_range` takes `paste` (`all`, `values`, `formulas`, `formats`, as in Paste Special),
  `transpose` (formula references swap their offsets as in Excel) and `skip_blanks`.
- `transform_range` runs `remove_duplicates` (chosen columns, header option, Excel's
  comparison: text ignores case, text and numbers differ), `text_to_columns` (delimiters,
  fixed widths, text qualifier, merged delimiters, Excel's type conversion) and `fill` (Fill
  Down and Right, and series: linear, growth and date with step and stop, as in
  Fill > Series).
- `replace_cells` replaces text in cell values and formulas, in one sheet or the whole
  workbook, matching case or whole cells. Results are retyped as in Excel (`1` becomes a
  number) and replaced formulas pass the formula check. `find_cells` stays read-only.
- The calculator evaluates `INDIRECT` (A1 and R1C1 text, names, `ROW(INDIRECT("1:10"))`),
  `CELL` (`address`, `row`, `col`, `contents`, `type`, `filename`, `prefix`, `protect`, and
  `width`, `format`, `color`, `parentheses` for cells without custom widths or number
  formats) and `INFO("recalc")`, checked against Excel's results. `CELL("filename")` shows the
  folder the way the server shows paths (relative to `--allow-dir`), not the host's path. Answers that depend on
  the host (`INFO("osversion")` and the like) are left uncalculated.

- `set_sheet_layout` can hide, show, group and ungroup rows and columns (`rows`, `columns`),
  hide or show a whole sheet (`visibility`; the last visible sheet cannot be hidden), set
  up printing (`print_setup`: orientation, paper size, scale or fit to pages, margins in cm,
  print area, repeated title rows and columns, centering, gridlines, header and footer) and
  protect or unprotect a sheet (`protection`: optional password, allowed actions).
- `insert_image` places a PNG or JPEG file at a cell, optionally sized in cm with the
  aspect ratio kept; `delete_image` removes one. `describe_sheet` lists the images and now
  also reports hidden rows and columns, the print area and whether the sheet is protected.
- `create_chart` takes either a `source` block (series in columns, or in rows with
  `series_in`) or explicit `series`: any ranges on any sheet (`'Sheet'!B2:B13`), not
  necessarily adjacent, with their own `name` (text, or a cell the name follows) and
  `categories`. Options that do not fit the chart type are rejected with an explanation.
- `scatter` charts have Excel's five subtypes (`scatter_style`: markers, lines with markers,
  lines, smooth with markers, smooth) and there are `bubble` charts (x, y and `sizes`).
- Series can have a `trendline` (linear, exponential, logarithmic, polynomial, power,
  moving average; optional equation and R²) and `error_bars` (fixed, percent, standard
  deviation or error; both, plus or minus; x or y).
- Axes (`x_axis`, `y_axis`, `secondary_y_axis`) take a `title`, `min`, `max`, `major_unit`,
  `log` scale, `reverse` order, `number_format`, major and minor gridlines, and the position
  of the tick `labels`. Any series can sit on the secondary axis (`secondary_axis`), and
  a series `type` (column, line, area) makes combo charts in either direction.
- Series formatting: `color`, `line_width_pt`, `marker` and `marker_size`, `data_labels`
  (which content, `position`, `number_format`), plus the chart's `colors`, `grouping`,
  `title_size`, `plot_color` and Excel 2007 `style` number.
- A chart can sit on its own chart sheet (omit `at`); `describe_workbook` lists
  `chart_sheets` and `delete_sheet` removes one.
- `create_chart` with `replace` (the name of a chart) replaces that chart in place, so a chart
  is changed by describing it again; `describe_sheet` lists each chart's `series` ranges.
- `create_chart` draws `doughnut` and `radar` charts.
- `describe_sheet` lists each chart with its `name`, `type`, `title` and `range`.
- `delete_chart` removes a chart by the name `describe_sheet` shows.
- `sort_range` sorts a range's rows by one or more columns (header text or column letter),
  ascending or descending, in Excel's order. Formatting, notes, links and formulas move
  with their rows.
- `set_defined_name` and `delete_defined_name` manage workbook- and sheet-scoped names
  for ranges and constants. The reference passes the formula safety check.
- `set_note` and `delete_note` manage cell notes; `describe_sheet` lists them.
- `create_pivot_table` adds a real Excel PivotTable (rows, columns, filters, and sum, count,
  average, min or max values, with subtotals and grand totals). Excel shows it at once,
  refreshes it from the source data and lets you rearrange it. The results are also written
  into the cells.
- `delete_pivot_table` removes a PivotTable and the cells it fills; `describe_sheet` lists
  each sheet's `pivot_tables` with their name, range and source.
- Opt-in VBA writing: start the server with `--allow-vba-write` (or
  `EXCEL_MCP_ALLOW_VBA_WRITE=1`) to get `write_vba_module` (set the code of a standard,
  class, `ThisWorkbook` or sheet module of an `.xlsm`/`.xltm`, creating the module or the
  whole VBA project if needed) and `delete_vba_module`. The tools are absent by default and
  in `--read-only` mode; code is stored, never run. `create_workbook` can then also create
  `.xlsm`/`.xltm` files.
- Editing a workbook through any tool keeps content that openpyxl cannot model, instead of
  deleting it: worksheet extensions (sparklines, Excel 2010 conditional formats and
  validation, slicer and timeline lists), newer chart types and their style parts (waterfall,
  histogram, treemap, sunburst, box and whisker, funnel), slicers and timelines with their
  caches, threaded comments and persons, cell and value metadata (dynamic arrays, linked data
  types, images in cells), shapes and text boxes, form controls with their VML, header and
  footer pictures, background pictures, custom XML, and the extension lists of PivotTables
  and their caches. It is read from the file when the workbook is opened for editing and put
  back, byte for byte, into the file openpyxl writes. Relationship ids, sheet and table
  numbers (which slicer caches use) and content types are kept consistent.
  `excel_mcp.package` documents the API that later features use to add such content.
- Content that would be left without what it refers to follows what Excel does: a slicer
  or timeline cache that no slicer uses is removed with its name, sparklines that read a
  deleted sheet are deleted, validation and conditional format formulas that refer to a
  deleted sheet become `#REF!`, threaded comments follow their notes when rows are inserted
  and go with a deleted note. `delete_sheet` and `delete_pivot_table` refuse when slicers
  would be left disconnected, which this server cannot store.
- Formulas that return several values are stored as dynamic array formulas, exactly as
  Excel stores them (an array formula over the spill range with the dynamic array metadata
  in `xl/metadata.xml`), with the results in the spilled cells. This covers `FILTER`, `SORT`,
  `SORTBY`, `UNIQUE`, `SEQUENCE`, `RANDARRAY`, `TRANSPOSE`, `TAKE`, `XLOOKUP` returning a
  range, operators and single-value functions applied to ranges (`=A2:A9*2`, `=SUM(B2:B9*C2:C9)`),
  and every other formula Excel itself saves that way. `write_range` reports formulas that a
  non-empty cell blocks in `blocked`; `read_range` shows spilled results, and the metadata
  survives later edits, moves with inserted rows, and is cleared with the formula.
- `read_range` in `formulas` mode returns the text of array formulas instead of an object
  description.

- `create_pivot_table` reproduces what Excel's PivotTable dialogs offer: `number_format` and
  `show_as` per values field (percent of total, row, column or parent, difference from, percent
  difference from, percent of, running total, percent of running total, rank, with
  `base_field` and `base_item`), `field_settings` per row, column or filter field (`show_items`,
  `sort` by label or by a values field with `sort_by`, `group_dates` for years,
  quarters, months and days, `group_numbers` for ranges), `calculated_fields` (formulas over
  the source fields, checked by the formula safety check), `layout` (`compact`, `outline`
  or `tabular`), `subtotals` on or off, and `values_in` `columns` or `rows`. Filters can show
  chosen items. The same field can be summarized twice (Excel's "Sum of Units2" naming). The
  figures in the cells are the ones Excel shows after refreshing the PivotTable, checked
  against about 200 PivotTables recorded from Excel.
- `write_range` takes `links`: it turns written cells into hyperlinks to `http`, `https` or
  `mailto` addresses or to places in the workbook (`#'Sheet 2'!A1`, a defined name), with a
  tooltip. Other kinds (`file:`, network shares, `javascript:`) are refused. `describe_sheet`
  lists `hyperlinks` and the sheet `view`. `clear_range` with `clear: "all"` removes a link.
- `set_sheet_layout` can move a sheet (`position`) and set how it looks when opened (`view`:
  `zoom`, `gridlines`, `headings`, `show_formulas`, `right_to_left`, `active`, `selected_cell`).
  `describe_workbook` marks the `active` sheet.
- `format_range` sets the `locked` and `formula_hidden` flags that decide what a protected
  sheet allows and shows.
- `set_workbook_settings` sets the document properties (title, subject, author, keywords,
  company), the calculation options (automatic, manual or automatic except data tables;
  iterative calculation with a maximum of iterations and change; recalculation on load) and
  workbook structure protection. `describe_workbook` returns them.
- The formula calculator reads `HYPERLINK` (its friendly name, or the link when there is none;
  checked against Excel) and sorts `SORT` and `SORTBY` over blank cells, logicals and errors in
  Excel's order: numbers, text, `FALSE`, `TRUE`, then errors by number, with blanks last
  whichever way it sorts. More than 200 cases recorded from Excel, including `UNIQUE` over
  blanks, logicals and errors, were added to its test fixture.

### Changed

- **Breaking:** `create_table` no longer takes `style` and `striped_rows`; they are fields of
  `options`.
- New tables use Excel's default style, `TableStyleMedium2`, instead of `TableStyleMedium9`.

- **Breaking:** `options.legend` defaults to Excel's own for the chart type instead of
  always the bottom: none for a single series (except pie and doughnut: bottom), bottom for
  several, top for radar with several series, waterfall and treemap, none for the other
  Excel 2016 charts. Explicit values still win. The Pareto percentage axis takes
  `secondary_y_axis` (title, min, max).
- **Breaking:** the conditional format field `icon_only` is now `hide_values`, and also
  applies to data bars.
- Rules of the Excel 2010 extension share their priority numbering with the classic ones.

- Invalid arguments are reported as one readable line per problem, such as
  `Invalid arguments for read_range: mode: 'formula' is not valid; use 'values' or
  'formulas'.`, instead of pydantic's report. Unknown fields get a suggestion
  (`rnage: unknown field; did you mean 'range'?`), several type failures of one value are
  merged, and input values are never echoed in full.
- Range errors say what is wrong (`Column ZZZ is past XFD, the last column.`), and messages
  that list names (sheets, fields, indices) share one format. An empty list reads
  `Sheet 'Data' has no images.` instead of `Valid image indices: none: ...`.
- Formulas that are not valid syntax fail with `InvalidFormulaError`; the formula safety
  policy still raises `UnsafeFormulaError`.
- `write_range` stores a formula that returns several values as a dynamic array formula.
  Excel used to show such a formula with `@` and a single value, because the file did not say
  it spills. Results of `=SUM(A1:A5*2)` and the like are now what Excel itself shows.
- `copy_sheet` carries the dynamic array metadata of the formulas it copies.
- `insert_rows_or_columns` and `delete_rows_or_columns` update references as Excel does,
  workbook-wide, instead of only moving cells. Formulas on every sheet (same-sheet and
  `Sheet!` references, absolute, relative and mixed, ranges that grow or shrink, whole
  rows and columns; a deleted cell becomes `#REF!`), defined names, conditional formats and
  data validation (including their relative formulas, which are split into parts when the
  edit cuts between a cell and what it refers to), merged cells, tables (columns, structured
  references), the sheet filter, print area and titles, freeze panes, row heights and column
  widths, page breaks, pictures, charts (on sheets and chart sheets, with their series,
  categories and titles), filters with their criteria and sort, and PivotTable locations and
  sources all follow the move, as do sparklines, extended conditional formats and validation,
  shapes, form controls and newer charts that Excel saved. Deleting one of several filtered
  columns applies the remaining criteria again, as Excel does. Like Excel, an edit fails (and nothing is saved) when it
  would cut through an array formula, a PivotTable, a table's header row or two tables at
  once. Inserted cells take the formatting of the line above or to the left, and rows
  inserted into a table get its calculated column formulas, as in Excel. Verified against
  Excel's own results for each edit (see the pull request).
- `rename_sheet` updates every reference to the sheet, as Excel does: formulas on all sheets
  (quoted only where needed), defined names, conditional formats, data validation, charts
  (also on chart sheets) and PivotTable sources.
- `INDIRECT`, `CELL` and `INFO` are allowed in formulas again: they only read the open
  workbook and the host, and the functions that could send data out stay blocked.
  `HYPERLINK` is allowed with a literal `http://`, `https://`, `mailto:` or `#Sheet!A1`
  link, and refused when its link is built from cells or points elsewhere.
- `Limits` gained `max_unpack_factor` and `max_compression_ratio` (uploads) and
  `max_copy_cells` (`copy_sheet`). An edit whose saved file would exceed the file size
  limit is refused and leaves the file unchanged.
- **Breaking:** `describe_workbook` returns `defined_names` as objects with `name`,
  `refers_to` and `sheet` (null for workbook scope), and includes sheet-scoped names.
- **Breaking:** `create_summary_table` is removed; `create_pivot_table` replaces it
  (`group_by` is now `row_fields`, `aggregation` is `function`, and the result is a PivotTable).
- `read_range`, `find_cells` and `describe_workbook` stream the workbook instead of loading
  it: memory stays flat (about 20 MB instead of 850 MB for a 200,000 x 10 sheet) and
  large files no longer risk exhausting memory. `describe_sheet` still loads the file.
- **Breaking:** compact results. `read_range` returns `{range, values, next_range?}`
  (no `sheet` or `truncated`; page on while `next_range` is present) and omits trailing
  empty cells and rows. `find_cells` returns `matches` grouped as `{sheet: {cell: value}}`.
  `describe_workbook` returns `sheets` (`name`, `used_range`, `visibility` when not visible),
  `defined_names` (when any) and `has_vba` (when true), without path, size, row and column counts.
  `describe_sheet` omits empty fields and its `name`.
  `list_workbooks` returns `{path: size_bytes}`. Dates at midnight read as `2026-01-31`.
- **Breaking:** consistent tool parameters and results, the last API change before 2.0. One
  name per concept: `sheet`, `range`, `at` (the top-left cell of whatever is placed or
  written), `source` (where data comes from, optionally `Data!A1:E200`) and `name` (of an
  object). Units are in the name. Charts and images are selected by name, which survives
  other deletions. No shims or aliases; behaviour is unchanged. Renames, old to new
  (many of these tools are new in 2.0, so the old names never shipped):
  - `create_sheet`: `sheet` to `new_name`, as in `rename_sheet` and `copy_sheet`.
  - `insert_rows_or_columns`, `delete_rows_or_columns`: `at` to `start`, so that `at` always
    is a cell.
  - `write_range`: `start_cell` to `at`.
  - `copy_range`: `target_cell` to `at`, `target_sheet` to `to_sheet`.
  - `create_pivot_table`: `source_sheet` and `source_range` to `source` (`Data!A1:E200`),
    `target_sheet` to `sheet`, `target_cell` to `at`, `rows` to `row_fields`, `columns` to
    `column_fields`, `values` to `value_fields`, `filters` to `filter_fields`, `fields` to
    `field_settings`.
  - `create_chart`: `anchor_cell` to `at`, `data_range` to `source`; `index` is replaced by
    `replace` (the name of the chart to replace) and `name` is new.
  - `delete_chart`, `delete_image`: `index` to `name`. `describe_sheet` lists each chart and
    image by `name` (instead of `index`) with the `range` of cells it covers (instead of
    `anchor`), and keeps the names Excel gave them, which edits no longer rewrite to
    `Chart 1`, `Image 2`.
  - `insert_image`: `cell` to `at`; `name` is new.
  - `add_slicer`: `cell` to `at`, `source` to `target`. `describe_sheet` lists slicers with
    `target` and `range` instead of `source` and `cell`.
  - `add_sparklines`: `location` to `range`, `data` to `source`.
  - `edit_table`: `table` to `name`.
  - `set_sheet_layout`: `column_widths` to `column_widths_chars`, `row_heights` to
    `row_heights_pt`. Chart series `line_width` and sparkline `line_weight` become
    `line_width_pt`. `describe_sheet` returns `column_widths_chars` as a list of
    `{column, width}`, `tables` as `{name, range}` and `hyperlinks` as `{cell, target,
    tooltip}` objects like its other collections.
  - `describe_workbook`: a sheet's `hidden` becomes `visibility` (`hidden` or `very_hidden`,
    omitted when visible), and `objects` counts its tables, charts, PivotTables, slicers and
    images.
- **Breaking:** tools that change a workbook return a small object instead of a sentence, with
  what the next call needs: `sheet`, `range`, `name` (and `path`, and a `note` where there
  is something to say). `create_chart` returns the chart's `name` and `range`,
  `create_pivot_table` and `create_table` their `name` and `range`, `insert_image` and
  `add_slicer` theirs; `insert_rows_or_columns` the `range` of the new lines (`3:4`).
- Tool descriptions say how `write_range` stores strings (text, unless `=` or an ISO date),
  that sheet protection does not stop this server's writes, and to leave a PivotTable's Grand
  Total out of a chart's source.
- Tool results are sent as compact JSON without default values, tools that return a message
  no longer advertise an output schema, and tool schemas lose generated titles and `null`
  unions: `tools/list` shrinks by about a third.
- `import_workbook` checks uploads without loading every cell.
- **Breaking:** `create_chart`'s `show_legend` option is replaced by `legend`: `bottom`
  (default, as in Excel), `right`, `left`, `top` or `none`.
- **Breaking:** `create_chart` loses `data_sheet` (put the sheet in the range, `Data!A1:C13`)
  and `secondary_line_columns` (give a series `type: "line"` and `secondary_axis`). Its
  `data_labels` option is an object (`{}` for values), `y_axis_min`, `y_axis_max`,
  `y_axis_number_format` and `x_axis_title`/`y_axis_title` move into `y_axis` and `x_axis`,
  and `at` is optional.
- Charts are drawn as current Excel draws them: gray text and light gridlines, no rounded
  corners, one color per series (not per category), a gap between clustered columns.
- **Breaking:** `describe_sheet` no longer returns `chart_count`; use the length of `charts`.
- Scatter charts plot points instead of joining them with lines.
- Horizontal bar charts list the rows top-down in sheet order, instead of Excel's
  bottom-up default.
- Line charts draw straight lines without markers unless `smooth` or `markers` is set;
  Excel used to curve them.
- Tool descriptions and input schemas are about 26% smaller (about 37,000 to 27,500
  characters across the 35 tools), without losing information a model needs: shorter
  parameter descriptions, no repetition between docstrings and fields, and `create_chart`
  no longer embeds the default options in its schema. The workbook directory is explained
  once in the server instructions instead of in every `path` description.
- **Breaking:** `set_sheet_layout` takes `column_widths_chars` and `row_heights_pt` as objects,
  `{"A": 20}` and `{"1": 30}`, instead of lists of `{column, width}` and `{row, height}`;
  this matches what `describe_sheet` returns for column widths.
- **Breaking:** `set_sheet_layout`'s `auto_filter` is an object (`range`, `filters`,
  `remove`) instead of a range string, and a range inside a table now points to the table's
  name instead of failing with "has its own filter".
- **Breaking:** `add_conditional_format` rules that format cells (`cell_value`, `formula` and
  the new types) need a `fill_color` or `font_color`, and fields that do not apply to the
  rule type are rejected. New rules take the next free priority instead of a count-based one.
- **Breaking:** `add_data_validation`'s `error_message` and `prompt` are limited to Excel's
  255 characters, and a list needs either `options` or `source`.
- **Breaking:** `create_table` and `edit_table` are one tool, `set_table(path, sheet, options,
  name, range)`: a table of that `name` on the sheet is changed (options, calculated columns,
  totals, resizing to `range`); otherwise `range` creates a new table (named `name`, or
  TableN). Behaviour and the table XML are unchanged.
- Tool descriptions and schemas are about 30% smaller (68,900 to 47,600 characters for the
  default tools), which every client pays for in context: shorter docstrings and parameter
  descriptions, no `additionalProperties: false` (unknown fields are still rejected by the
  server), no `default` for an absent boolean or list, no
  `null` option for optional parameters, plain type unions as one `type` list, and no
  `minimum` of 0 or 1. `path`, `sheet`, cell and range parameters have no description of their own; the
  server instructions explain paths and A1 notation. A test keeps the total under a budget.
- `create_chart` is marked destructive (it can replace a chart with `replace`), like the other
  tools that overwrite.
- `describe_workbook` leaves out hidden defined names (the `_xlchart` names of the newer
  chart types and other internal ones), as Excel's Name Manager does. Slicer names are listed.
- `create_workbook` describes `overwrite`, `list_workbooks` says that it lists the first
  workbook directory by default, and `read_range` says what `uncalculated` holds.

### Fixed

- `read_range` no longer loads the whole workbook (twice, about 1.5 GB and four times slower
  on 200,000 rows) when a formula returns empty text: Excel stores `""` as a result without
  content, which was mistaken for a formula never calculated. Only formulas with no stored
  result are calculated.
- Formulas are shown as typed, not as stored: `read_range` and `find_cells` in `formulas`
  mode, `replace_cells`, defined names and validation rules no longer show `_xlfn.`,
  `_xlws.` or `_xlpm.` prefixes, and spill references read `A1#`. Writing a shown formula
  back stores exactly what was stored.
- Results of dynamic array formulas (`FILTER`, `SORT`, `SEQUENCE`, ...) were stale after the
  data they use changed. Every save recalculates them, resizing the spill range as Excel does
  (`#SPILL!` where cells are in the way) and dropping a result the calculator cannot
  reproduce, which Excel calculates when it opens the file. Legacy (Ctrl+Shift+Enter) array
  formulas are recalculated into their range too. Every edit also sets
  `fullCalcOnLoad`, so Excel recalculates the whole workbook on open.
- `insert_rows_or_columns` and `delete_rows_or_columns` work through a dynamic array's spill
  range, which spills again, as in Excel. Legacy (Ctrl+Shift+Enter) array formulas still
  refuse it.
- `import_workbook` into an `.xlsx` or `.xltx` path removes the VBA project, as Excel does when
  saving without macros, instead of leaving an orphan `vbaProject.bin`.
- New workbooks no longer name "openpyxl" as their creator, and editing a file that has no
  creator does not add one.
- Dates written with `write_range` into a column of default width widen it to fit, as Excel
  does, instead of showing `####`.
- A drive-relative path such as `C:book.xlsx` gets a clear error.
- **Breaking:** tools that change a workbook no longer list an output schema in `tools/list`
  (about 5 KB less); they still return the same structured content. Only tools that return
  data (`read_range`, `describe_*`, `find_cells`, `list_workbooks`, `read_vba`) keep one.

- `copy_sheet` of a sheet with slicers or timelines (from Excel, or made here) wrote a slicer
  part that Excel could not open: the copied entry used the `xr10:` prefix without declaring
  it. Prefixes used by copied content are now declared on the root of the part it lands in
  and listed in `mc:Ignorable` as in the original, the copies get new `xr10:uid`s and new
  slicer names the way Excel numbers them ("Region 1" is copied as "Region 2", not
  "Region 1 1"), and the XML namespaces of the original's drawing and sheet are kept, so the
  newer chart anchors a copy made undeclared `xdr:` prefixes too.
- `copy_sheet` dropped shapes, text boxes, connectors, groups, picture fills, form controls,
  embedded (OLE) objects and protected ranges. Ignored errors are not copied, as in Excel. What
  cannot be copied (other sheet extensions, threaded comments) is named in the result's `note`.
  It copies them as Excel's "Create a copy" does: connectors stay attached to the copied
  shapes, controls get their own properties parts and ids, and what they read or set on the
  original they read or set on the copy. Shape ids that collide with ones openpyxl wrote are
  renumbered with the connectors that name them.
- `add_slicer` on a PivotTable with grouped dates failed with an opaque "Error executing tool".
  Slicers and timelines now work on grouped dates as in Excel, on the date field and on the
  groups made from it (`Years (Date)`, `Quarters (Date)`, `Months (Date)`, `Days (Date)`), and
  on fields grouped into number ranges: items outside a group's range are listed last and
  marked as having no data, number ranges and days are sorted as text, and the figures of
  PivotTables made here follow the selection (checked in Excel against its own files and
  refresh). `describe_sheet` lists a PivotTable's `date_groups`; a timeline on a field that
  holds no dates names the date fields.
- A workbook with several PivotTable caches saved every cache after the first linked to the
  first cache's records (an openpyxl quirk), so a slicer added after an earlier edit read the
  wrong records and wrote a field's shared items as indexes (`<x v="0"/>`) that Excel could
  not open. Each cache keeps its own records.
- A tool that fails unexpectedly now reports "Unexpected error in <tool>; see server log" and
  logs the exception, instead of the SDK's "Error executing tool <tool>".
- A parameter's own description was replaced by the one of its type in the schema (for
  example `at` and `source` of `create_chart` read "Cell" and "Range").

- `format_range` `font_name` took no effect in Excel on files this server creates: the font
  kept its theme font scheme, so Excel used the theme's font. Setting a name now clears the
  scheme, as picking a font in Excel does.
- CI: a stuck test now dumps every thread's stack after 60 seconds and fails after 120,
  and the test job times out after 8 minutes instead of 15.
- Editing a workbook no longer damages its charts: openpyxl dropped the chart style
  number, the rounded-corners flag, the plot area fill, the axes of area charts, and turned
  the empty text of chart labels into the word "None".
- Date and time limits in data validation (`2026-01-31`) were stored as the formula
  `2026-01-31`, a wrong number, instead of a date.
- Inserting or deleting rows or columns moves hyperlinks with their cells (they stayed in
  place before), and `copy_range` copies them.
- Editing a workbook no longer silently deletes sparklines, slicers, timelines, threaded
  comments, newer chart types, shapes, form controls and everything else listed under
  Added. Array formulas keep their range when rows are inserted above them.
- `copy_sheet` now copies what Excel's "Create a copy" does: data validation, conditional
  formats, images, charts (re-pointed at the copy's own data), tables (renamed, as Excel
  does), PivotTables, freeze panes, filters, print setup, protection and sheet-scoped
  names. It used to drop all of these. References to the sheet itself point at the copy.
- `set_sheet_layout` no longer writes overlapping column definitions when it changes a
  column that Excel stored together with its neighbours.
- Chart titles, axis titles and legends no longer sit on top of the plot in Excel.
- Formulas that use `IFS`, `XLOOKUP`, `TEXTJOIN`, `STDEV.S`, `SORT` and other newer
  functions showed `#NAME?` when the file was opened in Excel (and `SORT` made it
  unopenable) because the storage prefix was missing.
- `read_range` of a workbook whose formulas have no stored results (any file last saved by
  this server) calculates only the formulas in the range and the cells they use, reading
  rows of the sheet as they are needed, instead of loading the whole workbook twice. On
  200,000 rows with a formula in each, a page took 36.8 s and 1.1 GB, now 1.6 s and 230 MB;
  the used range 47.6 s, now 15 s (finding the used range is most of it); the last rows
  69.6 s, now 24 s (the sheet is read once to reach them).
- A whole-column or whole-row reference used where one value is expected now reads as
  blank past the last used cell in the calculator instead of `#VALUE!`.

### Security

- `import_workbook` checks every XML part of the package wherever it is stored, rejects
  packages with duplicate or unsafe part names, data connections, query tables or external
  links, caps the expanded size and the compression ratio (also while reading), and scans
  in constant memory with a nesting limit. It also rejects linked OLE objects and external
  relationships other than hyperlinks. Hyperlinks to files, network locations and other
  schemes than `http`, `https` and `mailto` (or with credentials) are removed instead of
  refusing the upload, and `import_workbook` reports in `note` which links were removed.
- Network (UNC) paths, Windows device paths and reserved device names are rejected for
  every file the server opens, with or without `--allow-dir`, unless the allowed folder is
  on that share. The rules are applied again to the resolved path, so a link that leads to a
  network share or device is refused too.
- Excel 4.0 macro functions that read files, the system or workbook internals, or run other
  programs (`FILES`, `GET.*`, `APP.*`, `RUN`, `EXEC`, `SEND.KEYS`, `ALERT`, `SQL.*`, `MAIL.*`
  and related) are blocked in every formula, including defined names. The list follows
  Microsoft's Excel 4.0 macro function index and was checked against Excel.
- The calculator matches wildcards without backtracking, limits formula length and nesting
  (64 levels, as in Excel), checks text and array sizes before building them, and ends
  hostile formulas as uncalculated cells instead of errors.
- Reading a VBA project fails fast on malformed containers and caps the module count and
  the total decompressed size.
- `write_vba_module` checks for attribute lines after normalising line endings.

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

[Unreleased]: https://github.com/haris-musa/excel-mcp-server/compare/v1.1.2...HEAD
[1.1.2]: https://github.com/haris-musa/excel-mcp-server/compare/v1.1.1...v1.1.2
[1.1.1]: https://github.com/haris-musa/excel-mcp-server/compare/v1.1.0...v1.1.1
[1.1.0]: https://github.com/haris-musa/excel-mcp-server/compare/v1.0.0...v1.1.0
[1.0.0]: https://github.com/haris-musa/excel-mcp-server/compare/v0.1.8...v1.0.0
[0.1.8]: https://github.com/haris-musa/excel-mcp-server/releases/tag/v0.1.8
