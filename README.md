# Excel MCP Server

A Model Context Protocol (MCP) server implementation that provides Excel file manipulation capabilities without requiring Microsoft Excel installation. This server enables workbook creation, data manipulation, formatting, and advanced Excel features.

## Requirements

- Python 3.10+
- MCP SDK 1.2.0+
- OpenPyXL 3.1.2+

## Components

### Resources

The server provides Excel workbook manipulation through OpenPyXL:

- Creates and modifies Excel workbooks
- Manages worksheets and ranges
- Handles formatting and styles
- Supports charts and pivot tables

### Tools

This server provides a comprehensive set of Excel manipulation tools. For detailed documentation of all available tools, their parameters, and usage examples, please refer to [TOOLS.md](TOOLS.md).

The tools include capabilities for:

- Workbook and worksheet management
- Data reading and writing
- Formatting and styling
- Charts and visualizations
- Pivot tables and data analysis

See [TOOLS.md](TOOLS.md) for complete documentation.

## Features

- Full Excel Support: Comprehensive Excel functionality
- Data Manipulation: Read, write, and transform data
- Advanced Features: Charts, pivot tables, and formatting
- Error Handling: Comprehensive error handling with clear messages

## Usage

### Environment Configuration

The server can be configured using the following environment variables:

- `EXCEL_FILES_PATH`: Directory where Excel files will be stored (default: `./excel_files`)

You can set this in different ways:

Windows CMD:

```cmd
set EXCEL_FILES_PATH=C:\path\to\excel\files
uv run excel-mcp-server
```

Windows PowerShell:

```powershell
$env:EXCEL_FILES_PATH="C:\path\to\excel\files"
uv run excel-mcp-server
```

Linux/MacOS:

```bash
export EXCEL_FILES_PATH=/path/to/excel/files
uv run excel-mcp-server
```

Or in Claude Desktop config:

```json
{
  "mcpServers": {
    "excel": {
      "command": "uv run excel-mcp-server",
      "transport": "sse",
      "env": {
        "EXCEL_FILES_PATH": "/path/to/excel/files"
      }
    }
  }
}
```

### Starting the Server

Start the server:

```bash
uv run excel-mcp-server
```

The server will start in SSE mode and wait for connections from MCP clients.

### Connecting in Cursor IDE

After starting the server, connect to the SSE endpoint in Cursor IDE:

```
http://localhost:8000/sse
```

The Excel MCP tools will be available through the agent.

For available tools and their usage, please refer to [TOOLS.md](TOOLS.md).

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
