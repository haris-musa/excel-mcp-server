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
  50 megapixels, and are validated with Pillow before they are embedded.
- **Formulas**: every formula is tokenized and rejected if it uses a function that reaches
  the network, other programs or host information, a DDE link, or another workbook. A
  reference can only name sheets of the workbook it is in. Uploaded workbooks are checked
  the same way, including defined names (also when created by `set_defined_name`), conditional formats, data validation, tables and
  chart references. Formulas written by `copy_range`, `transform_range` (text to columns) and
  `replace_cells` pass the same check, as do list sources and rule operands. Formulas that `insert_rows_or_columns`,
  `delete_rows_or_columns` and `rename_sheet` rewrite pass the same check.
- **Calculator**: `read_range` evaluates formulas with a built-in interpreter that has no
  access to files, the network or other programs. It never evaluates the functions the
  formula check rejects, and its work (cells evaluated, nesting depth) is bounded per call.
- **Network**: HTTP binds to localhost by default with DNS rebinding protection. A public
  host requires `EXCEL_MCP_AUTH_TOKEN` unless `--allow-unauthenticated` is passed for a
  deployment where a proxy handles authentication.
- **Resources**: limits on file size and on the cells read or written per call.
- **Preserved content**: parts of a workbook that the server cannot edit (extensions,
  slicers, newer charts, comments, form controls, custom XML, embedded objects and data
  connections that are part of the file) are carried over unchanged when it is edited, as
  the file already contained them. They are never run or fetched. Only their structure is
  read (XML with a DOCTYPE declaration is refused), and the amount of such content is
  limited by the file size limit, which also bounds a compressed bomb. Thumbnails and
  digital signatures are dropped, because they no longer match an edited file.
- **Macros**: VBA code is parsed as text by the server's own [MS-OVBA] reader with size
  limits, and is never run. Macros in `.xlsm` files are kept intact. Writing macros is off by
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
