# Running Excel MCP Server in Docker

This document provides instructions on how to run Excel MCP Server using Docker.

## Using Docker

### Prerequisites

- Install [Docker](https://docs.docker.com/get-docker/)
- Install [Docker Compose](https://docs.docker.com/compose/install/) (optional, but recommended)

## New Smart Start Method (Recommended)

We provide a smart startup script that can automatically read the Excel file path you set in the Cursor IDE's mcp.json configuration and correctly configure the Docker container. This way, you can modify the path in Cursor IDE at any time without having to modify the Docker configuration.

### Using the Smart Startup Script

1. Make sure your Cursor IDE has mcp.json correctly configured:

```json
{
  "mcpServers": {
    "excel-mcp-server": {
      "url": "http://localhost:8000/sse",
      "env": {
        "EXCEL_FILES_PATH": "/your/chosen/path"
      }
    }
  }
}
```

2. Run the startup script:

```bash
./start-excel-mcp.sh
```

The script will automatically:
- Read the `EXCEL_FILES_PATH` you configured in mcp.json
- Ensure that directory exists and has correct permissions
- Configure the Docker container to use that path
- Start the service

3. View logs:

```bash
docker-compose logs -f
```

4. Stop the service:

```bash
docker-compose down
```

### Notes on Path Mounting

This Docker configuration uses a flexible path mounting strategy:

1. Automatically mounts the path read from mcp.json
2. Default mounts for common user directory paths:
   - `/Users` directory (macOS)
   - Current user's home directory (`$HOME`)
   - Relative path `./excel_files` mounted to `/app/excel_files` inside the container

### Docker Desktop File Sharing Settings

**Important:** When using Docker Desktop on macOS, you must ensure that the directory you want to access has been added to Docker Desktop's file sharing settings:

1. Open Docker Desktop
2. Click the gear icon (⚙️) in the top-right corner to open settings
3. Select "Resources" > "File sharing"
4. Make sure the directory you need to access is added to the list
5. If not, click "+" to add the directory, then click "Apply & Restart"

## Manual Configuration Method (Not Recommended)

If you don't want to use the startup script, you can also manually configure and start the service:

### Method 1: Manually Set Environment Variables

```bash
# Set environment variables
export EXCEL_FILES_PATH="/your/path"

# Start the service
docker-compose up -d
```

### Method 2: Using Docker Command

```bash
docker run -d \
  --name excel-mcp-server \
  -p 8000:8000 \
  -v /Users:/Users \
  -v $HOME:$HOME \
  -v your/path:your/path \
  -e EXCEL_FILES_PATH=your/path \
  excel-mcp-server
```

## Configuration and Usage

After starting the service, the MCP server will run at http://localhost:8000/sse.

### Using in Cursor IDE

Add to your Cursor IDE configuration (each user can set their own path):

```json
{
  "mcpServers": {
    "excel-mcp-server": {
      "url": "http://localhost:8000/sse",
      "env": {
        "EXCEL_FILES_PATH": "/your/chosen/path"
      }
    }
  }
}
```

**Important Notes:** 
- macOS users should choose a path that starts with `/Users` or is within their home directory
- Make sure the chosen path has been added to Docker Desktop's file sharing settings
- After modifying the path in mcp.json, you need to run the startup script again to make the changes effective

## Troubleshooting

1. If port 8000 is already in use, you can modify the port mapping in `docker-compose.yml`:

```yaml
ports:
  - "8001:8000"  # Map port 8001 on the host to port 8000 in the container
```

2. If you encounter permission issues, you may need to adjust the permissions of your Excel files directory:

```bash
chmod 755 /your/path
```

3. If the Docker container cannot access the mounted directory, which is the most common problem, make sure:
   - The directory has been added to Docker Desktop's file sharing settings
   - On macOS, only limited directories like `/Users`, `/Volumes`, `/private`, `/tmp` can be accessed
   - Check Docker Desktop's Settings > Resources > File sharing section

4. The "Mounts denied" error usually indicates that Docker does not have permission to access the specified path. Solution:
   - Open Docker Desktop settings
   - Select "Resources" > "File sharing"
   - Add the directory you need to access
   - Apply changes and restart Docker

5. If the startup script cannot correctly read mcp.json, make sure:
   - mcp.json is located at `~/.cursor/mcp.json`
   - The file format is correct
   - Consider installing the jq tool (`brew install jq`) for better JSON parsing support