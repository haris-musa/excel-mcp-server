<p align="center">
  <img src="https://raw.githubusercontent.com/haris-musa/excel-mcp-server/main/assets/logo.png" alt="Excel MCP Server Logo" width="300"/>
</p>

[![PyPI version](https://img.shields.io/pypi/v/excel-mcp-server.svg)](https://pypi.org/project/excel-mcp-server/)
[![Total Downloads](https://static.pepy.tech/badge/excel-mcp-server)](https://pepy.tech/project/excel-mcp-server)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![smithery badge](https://smithery.ai/badge/@haris-musa/excel-mcp-server)](https://smithery.ai/server/@haris-musa/excel-mcp-server)
[![Install MCP Server](https://cursor.com/deeplink/mcp-install-dark.svg)](https://cursor.com/install-mcp?name=excel-mcp-server&config=eyJjb21tYW5kIjoidXZ4IGV4Y2VsLW1jcC1zZXJ2ZXIgc3RkaW8ifQ%3D%3D)

A Model Context Protocol (MCP) server that lets you manipulate Excel files without needing Microsoft Excel installed. Create, read, and modify Excel workbooks with your AI agent.

## Features

- 📊 **Excel Operations**: Create, read, update workbooks and worksheets
- 📈 **Data Manipulation**: Formulas, formatting, charts, pivot tables, and Excel tables
- 🔍 **Data Validation**: Built-in validation for ranges, formulas, and data integrity
- 🎨 **Formatting**: Font styling, colors, borders, alignment, and conditional formatting
- 📋 **Table Operations**: Create and manage Excel tables with custom styling
- 📊 **Chart Creation**: Generate various chart types (line, bar, pie, scatter, etc.)
- 🔄 **Pivot Tables**: Create dynamic pivot tables for data analysis
- 🔧 **Sheet Management**: Copy, rename, delete worksheets with ease
- 🔌 **Triple transport support**: stdio, SSE (deprecated), and streamable HTTP
- 🌐 **Remote & Local**: Works both locally and as a remote service

## Usage

The server supports multiple deployment methods:

### 🚀 Docker Deployment (Recommended)

The easiest way to deploy the Excel MCP Server is using Docker. We provide convenient startup scripts for both development and production environments.

#### Quick Start

1. **Copy environment configuration**:
   ```bash
   cp env.example .env
   ```

2. **Edit configuration (optional)**:
   ```bash
   # Edit .env file to configure your Excel file path and port
   # HOST_EXCEL_PATH=./excel_files  # Excel file storage path
   # HOST_PORT=8000                 # External service port
   ```

3. **Start the service**:

   **Linux/macOS:**
   ```bash
   chmod +x start.sh
   ./start.sh start
   ```

   **Windows PowerShell:**
   ```powershell
   .\start.ps1 start
   ```

#### Docker Management Commands

**Using startup scripts:**
```bash
# Start service
./start.sh start

# Stop service
./start.sh stop

# Restart service
./start.sh restart

# Check status
./start.sh status

# View logs
./start.sh logs
```

**Using Docker Compose directly:**
```bash
# Start service (development)
docker-compose up -d

# Start service (production)
docker-compose -f docker-compose.prod.yml up -d

# Stop service
docker-compose down

# View logs
docker-compose logs -f excel-mcp-server

# Rebuild image
docker-compose build --no-cache
```

#### Docker Client Configuration

```json
{
   "mcpServers": {
      "excel": {
         "url": "http://localhost:8000/mcp"
      }
   }
}
```

### 📦 Direct Installation

#### 1. Stdio Transport (for local use)

```bash
uvx excel-mcp-server stdio
```

```json
{
   "mcpServers": {
      "excel": {
         "command": "uvx",
         "args": ["excel-mcp-server", "stdio"]
      }
   }
}
```

#### 2. SSE Transport (Server-Sent Events - Deprecated)

```bash
uvx excel-mcp-server sse
```

**SSE transport connection**:
```json
{
   "mcpServers": {
      "excel": {
         "url": "http://localhost:8000/sse",
      }
   }
}
```

#### 3. Streamable HTTP Transport (Recommended for remote connections)

```bash
uvx excel-mcp-server streamable-http
```

**Streamable HTTP transport connection**:
```json
{
   "mcpServers": {
      "excel": {
         "url": "http://localhost:8000/mcp",
      }
   }
}
```

## Environment Variables & Configuration

### 🐳 Docker Environment Variables

When using Docker deployment, configure these variables in your `.env` file:

| Variable | Description | Default Value | Notes |
|----------|-------------|---------------|-------|
| `HOST_EXCEL_PATH` | Host machine Excel file storage path | `./excel_files` | Can be absolute or relative path |
| `HOST_LOG_PATH` | Host machine log file storage path | `./logs` | For persistent container logs |
| `HOST_PORT` | Host machine external port | `8000` | Access via `http://localhost:<HOST_PORT>/mcp` |
| `FASTMCP_PORT` | Container internal service port | `8000` | Usually no need to modify |

#### Environment Configuration Example

```env
# Excel MCP Server Configuration

# Host machine Excel file storage path
# Can be absolute path like /home/user/excel_files
HOST_EXCEL_PATH=./excel_files

# Host machine log file storage path (optional)
HOST_LOG_PATH=./logs

# Host machine external port
HOST_PORT=8000

# Container internal service port
FASTMCP_PORT=8000
```

### 📦 Direct Installation Environment Variables

#### SSE and Streamable HTTP Transports

When running the server with the **SSE or Streamable HTTP protocols**, you **must set the `EXCEL_FILES_PATH` environment variable on the server side**. This variable tells the server where to read and write Excel files.
- If not set, it defaults to `./excel_files`.

You can also set the `FASTMCP_PORT` environment variable to control the port the server listens on (default is `8000` if not set).
- Example (Windows PowerShell):
  ```powershell
  $env:EXCEL_FILES_PATH="E:\MyExcelFiles"
  $env:FASTMCP_PORT="8007"
  uvx excel-mcp-server streamable-http
  ```
- Example (Linux/macOS):
  ```bash
  EXCEL_FILES_PATH=/path/to/excel_files FASTMCP_PORT=8007 uvx excel-mcp-server streamable-http
  ```

#### Stdio Transport

When using the **stdio protocol**, the file path is provided with each tool call, so you do **not** need to set `EXCEL_FILES_PATH` on the server. The server will use the path sent by the client for each operation.

## 📁 Project Structure

```
excel-mcp-server/
├── Dockerfile              # Docker image build file
├── docker-compose.yml      # Development environment configuration
├── docker-compose.prod.yml # Production environment configuration
├── env.example            # Environment variables template
├── start.sh               # Linux/macOS startup script
├── start.ps1              # Windows PowerShell startup script
├── src/                   # Source code
│   └── excel_mcp/         # Main package
└── TOOLS.md               # Available tools documentation
```

## 🔍 Verification & Testing

After deployment, verify that the service is running correctly:

```bash
# Check MCP endpoint
curl -X POST -H "Content-Type: application/json" \
  -H "Accept: application/json, text/event-stream" \
  -d '{"jsonrpc": "2.0", "method": "initialize", "params": {"protocolVersion": "2024-11-05", "capabilities": {}}, "id": 1}' \
  http://localhost:8000/mcp/

# Check service status
docker-compose ps

# View logs
docker-compose logs -f excel-mcp-server
```

## 🔧 Troubleshooting

### Common Issues

#### 1. Port Conflicts
**Problem**: Port 8000 is already in use
**Solution**: 
- Modify `HOST_PORT` in `.env` file
- Restart service: `docker-compose down && docker-compose up -d`
- Update client configuration with new port

#### 2. Excel File Access Permissions
**Problem**: Cannot read/write Excel files
**Solution**:
- Ensure `HOST_EXCEL_PATH` directory has read/write permissions
- Check if directory exists: `ls -la ./excel_files`
- For Linux/macOS: `chmod 755 ./excel_files`

#### 3. Service Startup Failure
**Problem**: Container keeps restarting
**Solution**:
- Check logs: `docker-compose logs excel-mcp-server`
- Verify port availability: `netstat -tlnp | grep 8000`
- Ensure Docker daemon is running

#### 4. Health Check Failures
**Problem**: Container shows as unhealthy
**Solution**:
- Check if service is responding: `curl http://localhost:8000/mcp`
- Verify port mapping is correct
- Review container logs for errors

#### 5. Client Connection Issues
**Problem**: MCP client cannot connect
**Solution**:
- Verify service is running: `docker-compose ps`
- Check firewall settings
- Ensure correct URL in client configuration
- Test with: `curl -X POST http://localhost:8000/mcp`

### Log Analysis

```bash
# View container logs
docker-compose logs -f excel-mcp-server

# View host machine logs (if configured)
tail -f logs/excel-mcp.log

# Check Docker system logs
docker system events

# Debug container
docker-compose exec excel-mcp-server sh
```

## 🚀 Production Deployment

### Security Considerations

1. **Reverse Proxy Setup**
   - Use Nginx or Apache as reverse proxy
   - Enable HTTPS/SSL certificates
   - Configure rate limiting

2. **Access Control**
   - Implement authentication if needed
   - Use firewall rules to restrict access
   - Consider VPN for internal services

3. **Environment Variables**
   - Never commit `.env` file to version control
   - Use Docker secrets for sensitive data
   - Set appropriate file permissions (600) for `.env`

### Production Configuration

Use `docker-compose.prod.yml` for production deployment:

```bash
# Production deployment
docker-compose -f docker-compose.prod.yml up -d

# With custom environment
docker-compose -f docker-compose.prod.yml --env-file .env.prod up -d
```

**Key production differences:**
- Resource limits (CPU: 1.0 core, Memory: 512MB)
- Security options (no-new-privileges, specific user)
- Longer health check intervals
- Always restart policy
- System paths for logs and data

### Data Persistence

1. **Excel Files**
   - Use absolute paths for production: `/var/lib/excel-mcp/excel_files`
   - Regular backups of Excel files
   - Consider using external storage (NFS, S3)

2. **Logs**
   - Configure log rotation
   - Use centralized logging (ELK stack, Fluentd)
   - Set up log monitoring and alerting

3. **Database**
   - If using external databases, ensure connections are secure
   - Use connection pooling
   - Configure proper backup strategies

### Monitoring & Maintenance

1. **Health Monitoring**
   - Set up monitoring tools (Prometheus, Grafana)
   - Configure alerting for service downtime
   - Monitor resource usage

2. **Updates**
   - Plan regular updates for security patches
   - Test updates in staging environment first
   - Use blue-green deployment for zero downtime

3. **Scaling**
   - Configure multiple replicas if needed
   - Use load balancer for high availability
   - Consider container orchestration (Kubernetes)

## Available Tools

The server provides a comprehensive set of Excel manipulation tools. See [TOOLS.md](TOOLS.md) for complete documentation of all available tools.

## Star History

[![Star History Chart](https://api.star-history.com/svg?repos=haris-musa/excel-mcp-server&type=Date)](https://www.star-history.com/#haris-musa/excel-mcp-server&Date)

## License

MIT License - see [LICENSE](LICENSE) for details.
