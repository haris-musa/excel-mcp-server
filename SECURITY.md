# Security policy

## Reporting a vulnerability

Please report vulnerabilities privately through
[GitHub security advisories](https://github.com/haris-musa/excel-mcp-server/security/advisories/new),
not in public issues. Include the version, the transport (stdio or Streamable HTTP), and
steps or a proof of concept.

You can expect a first response within a week. Fixes are released as soon as they are
ready, and reporters are credited in the advisory and the changelog unless they prefer
otherwise.

## Supported versions

Security fixes are made for the latest release.

## Security model

The server acts on behalf of whoever can call its tools: the local MCP client in stdio mode,
or anyone who can reach the HTTP endpoint. Its defences are:

- **Files**: only Excel files (`.xlsx`, `.xlsm`, `.xltx`, `.xltm`), confined to the
  `--allow-dir` folders when they are set (always in HTTP mode), after resolving symlinks.
  Existing files are only replaced when a tool is asked to overwrite them. Pictures for
  `insert_image` follow the same folder rules, must be PNG or JPEG files up to 10 MB and
  50 megapixels, and are validated with Pillow before they are embedded. Network (UNC)
  paths, Windows device paths, reserved device names (`CON`, `NUL`, `COM1`, ...) and NTFS
  alternate data streams are rejected before anything is opened, with or without
  `--allow-dir`, because opening a network path can send the user's Windows credentials
  to that host. The one exception is a share that holds an allowed folder. A workbook is
  never saved larger than the file size limit.
- **Formulas**: every formula is tokenized and rejected if it uses a function that reaches
  the network or other programs, an Excel 4.0 macro function (`FILES`, `GET.WORKBOOK`,
  `RUN`, `EXEC`, ... which work in defined names), a DDE link, or another workbook. A
  reference can only name sheets of the workbook it is in. `HYPERLINK` is allowed only with
  a literal `http://`, `https://`, `mailto:` or `#Sheet!A1` link, since a link built from
  cells could carry their contents to another server when clicked. `INDIRECT`, `CELL` and
  `INFO` are allowed: they read the open workbook and the host but, with every function
  that sends data out blocked, cannot leak it. The aim is to block network, code execution,
  file system and resource-exhaustion channels and nothing else.
  Formulas written by `copy_range`, `transform_range` (text to columns) and
  `replace_cells` pass the same check, as do list sources and rule operands, and the
  formulas that `insert_rows_or_columns`, `delete_rows_or_columns` and `rename_sheet`
  rewrite.
- **Uploads**: `import_workbook` scans every XML part of the package, found by content and not
  by name, in constant memory, including defined names, conditional formats, data validation,
  tables and chart references. It rejects packages with duplicate or traversing part names,
  data connections, query tables, external workbook links, linked OLE objects and any other
  external relationship that is not a hyperlink, and packages that expand beyond
  `max_unpack_factor` times the file size limit or compress more densely than
  `max_compression_ratio`, both checked while reading. The same limits apply to every
  workbook the server opens, so a zip bomb already on disk is refused before it is read. Hyperlinks are not refused but
  neutralised: external links to anything other than `http`, `https` and `mailto` addresses
  without credentials (file paths, network shares, `smb:`, `ms-excel:`, `javascript:` and so
  on) are removed from cells, shapes, pictures and legacy drawings, the cell text and format
  stay, and the import result lists what was removed. `write_range` still refuses such links.
- **Calculator**: `read_range` evaluates formulas with a built-in interpreter that has no
  access to files, the network or other programs. It never evaluates the functions the
  formula check rejects, and its work (cells evaluated, nesting depth, text and array
  sizes) is bounded per call. Wildcards match in linear time, and formulas are limited to
  Excel's length and 64 nesting levels.
- **Network**: HTTP binds to localhost by default with DNS rebinding protection. A public
  host requires `EXCEL_MCP_AUTH_TOKEN` unless `--allow-unauthenticated` is passed for a
  deployment where a proxy handles authentication.
- **Resources**: limits on file size, on the cells read or written per call, and on the
  cells of a sheet that `copy_sheet` duplicates.
- **Preserved content**: parts of a workbook that the server cannot edit (extensions,
  slicers, newer charts, comments, form controls, custom XML, embedded objects and data
  connections that are part of the file) are carried over unchanged when it is edited, as
  the file already contained them. They are never run or fetched. Only their structure is
  read (XML with a DOCTYPE declaration is refused), and the amount of such content is
  limited by the file size limit, which also bounds a compressed bomb. Thumbnails and
  digital signatures are dropped, because they no longer match an edited file.
- **Macros**: VBA code is parsed as text by the server's own [MS-OVBA] reader, which bounds
  the module count and the total decompressed size and fails fast on malformed containers,
  and is never run. Macros in `.xlsm` files are kept intact. Writing macros is off by
  default: the `write_vba_module` and `delete_vba_module` tools exist only when the server is
  started with `--allow-vba-write` (`EXCEL_MCP_ALLOW_VBA_WRITE=1`), never in `--read-only`
  mode, and are marked destructive. They write only to `.xlsm`/`.xltm` files inside the
  allowed folders, limit the size of code, and only store it: whether macros run is up to
  Excel and the user. A prompt-injected model could plant harmful code with this option on,
  so users must review macro code before enabling macros.

Out of scope: an assistant acting on instructions hidden in cell text (prompt injection)
is a risk of the client and model; the server marks cell contents as untrusted data but
cannot prevent it. Run the server with `--read-only` or a narrow `--allow-dir` when working
with untrusted workbooks.
