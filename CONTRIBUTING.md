# Contributing

Thanks for helping. Bug reports, fixes and focused improvements are welcome.

## Setup

Install [uv](https://docs.astral.sh/uv/), then:

```bash
git clone https://github.com/haris-musa/excel-mcp-server
cd excel-mcp-server
uv sync
```

## Checks

Run these before opening a pull request; CI runs the same:

```bash
uv run ruff format .
uv run ruff check .
uv run pyright
uv run pytest
```

If you change a tool, regenerate the tool reference and update `manifest.json` if a tool
was added, removed or renamed:

```bash
uv run python scripts/generate_tools_doc.py
```

To try the server in a real client, use the
[MCP Inspector](https://github.com/modelcontextprotocol/inspector):

```bash
npx @modelcontextprotocol/inspector uv run excel-mcp-server stdio --allow-dir ./excel_files
```

## Guidelines

- Keep changes small and focused, with a test for every fix.
- Follow the project rules in [AGENTS.md](AGENTS.md): layering, security invariants, tool
  design and code style.
- Add a line to the `Unreleased` section of [CHANGELOG.md](CHANGELOG.md) for user-visible
  changes.
- Use [Conventional Commits](https://www.conventionalcommits.org/) for commit messages.

Report security problems privately as described in [SECURITY.md](SECURITY.md).
