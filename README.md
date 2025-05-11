# Excel MCP Server

A Model Context Protocol (MCP) server that lets you manipulate Excel files without needing Microsoft Excel installed. Create, read, and modify Excel workbooks with your AI agent.

## Features

- 📊 Create and modify Excel workbooks
- 📝 Read and write data
- 🎨 Apply formatting and styles
- 📈 Create charts and visualizations
- 📊 Generate pivot tables
- 🔄 Manage worksheets and ranges
- 🔌 Dual transport support: stdio and SSE

## Quick Start

### Prerequisites

- Python 3.10 or higher

### Installation

1. Clone the repository:
```bash
git clone https://github.com/haris-musa/excel-mcp-server.git
cd excel-mcp-server
```

2. Install using uv:
```bash
uv pip install -e .
```

### Running the Server

The server supports two transport modes: stdio and SSE.

#### Using stdio transport (default)

Stdio transport is ideal for direct integration with tools like Cursor Desktop or local development, which can manipulate local files:

```bash
excel-mcp-server stdio
```

#### Using SSE transport

SSE transport is perfect for web-based applications and remote connections, which manipulate remote files:

```bash
excel-mcp sse
```

You can specify host and port for the SSE server:

```bash
excel-mcp sse --host 127.0.0.1 --port 8080
```

## Using with AI Tools

### Cursor IDE

1. Add this configuration to Cursor, choosing the appropriate transport method for your needs:

**Stdio transport connection** (for local integration):
```json
{
   "mcpServers": {
      "excel-stdio": {
         "command": "uv",
         "args": ["run", "excel-mcp-server", "stdio"]
      }
   }
}
```

**SSE transport connection** (for web-based applications):
```json
{
   "mcpServers": {
      "excel": {
         "url": "http://localhost:8000/sse",
         "env": {
            "EXCEL_FILES_PATH": "/path/to/excel/files"
         }
      }
   }
}
```

2. The Excel tools will be available through your AI assistant.

### Remote Hosting & Transport Protocols

This server supports both stdio and SSE transport protocols for maximum flexibility:

1. **Using with Claude Desktop:**
   - Use Stdio transport

2. **Hosting Your MCP Server (SSE):**
   - [Remote MCP Server Guide](https://developers.cloudflare.com/agents/guides/remote-mcp-server/)

## Environment Variables （for SSE transport）

- `FASTMCP_PORT`: Server port for SSE transport (default: 8000)
- `EXCEL_FILES_PATH`: Directory for Excel files (default: `./excel_files`)

## Available Tools

The server provides a comprehensive set of Excel manipulation tools. See [TOOLS.md](TOOLS.md) for complete documentation of all available tools.

## License

MIT License - see [LICENSE](LICENSE) for details.
