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
  Existing files are only replaced when a tool is asked to overwrite them.
- **Formulas**: every formula is tokenized and rejected if it uses a function that reaches
  the network, other programs or host information, a DDE link, or another workbook.
- **Network**: HTTP binds to localhost by default with DNS rebinding protection. A public
  host requires `EXCEL_MCP_AUTH_TOKEN` unless `--allow-unauthenticated` is passed for a
  deployment where a proxy handles authentication.
- **Resources**: limits on file size and on the cells read or written per call.

Out of scope: an assistant acting on instructions hidden in cell text (prompt injection)
is a risk of the client and model; the server marks cell contents as untrusted data but
cannot prevent it. Run the server with `--read-only` or a narrow `--allow-dir` when working
with untrusted workbooks.
