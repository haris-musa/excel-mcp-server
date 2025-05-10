# 在Docker中运行Excel MCP Server

这个文档提供了如何使用Docker来运行Excel MCP Server的说明。

## 使用Docker

### 前提条件

- 安装 [Docker](https://docs.docker.com/get-docker/)
- 安装 [Docker Compose](https://docs.docker.com/compose/install/) (可选，但推荐)

## 新的智能启动方式 (推荐)

我们提供了一个智能启动脚本，它能够自动从Cursor IDE的mcp.json配置中读取您设置的Excel文件路径，并正确配置Docker容器。这样，您可以随时在Cursor IDE中修改路径，而不需要修改Docker配置。

### 使用智能启动脚本

1. 确保您的Cursor IDE中已正确配置mcp.json:

```json
{
  "mcpServers": {
    "excel-mcp-server": {
      "url": "http://localhost:8000/sse",
      "env": {
        "EXCEL_FILES_PATH": "/您选择的路径"
      }
    }
  }
}
```

2. 运行启动脚本:

```bash
./start-excel-mcp.sh
```

脚本会自动:
- 从mcp.json中读取您配置的`EXCEL_FILES_PATH`
- 确保该目录存在并有正确的权限
- 配置Docker容器使用该路径
- 启动服务

3. 查看日志:

```bash
docker-compose logs -f
```

4. 停止服务:

```bash
docker-compose down
```

### 关于路径挂载的说明

此Docker配置采用了灵活的路径挂载策略:

1. 自动挂载从mcp.json读取的路径
2. 默认挂载了常见的用户目录路径:
   - `/Users` 目录 (macOS)
   - 当前用户的主目录 (`$HOME`)
   - 相对路径 `./excel_files` 挂载到容器内的 `/app/excel_files`

### Docker Desktop文件共享设置

**重要：** 在macOS上使用Docker Desktop时，您必须确保要访问的目录已在Docker Desktop的文件共享设置中添加：

1. 打开Docker Desktop
2. 点击右上角的齿轮图标（⚙️）打开设置
3. 选择"Resources" > "File sharing"
4. 确保您需要访问的目录已添加到列表中
5. 如果没有，点击"+"添加目录，然后点击"Apply & Restart"

## 手动配置方式 (不推荐)

如果您不想使用启动脚本，也可以手动配置和启动服务:

### 方法一：手动设置环境变量

```bash
# 设置环境变量
export EXCEL_FILES_PATH="/您的路径"

# 启动服务
docker-compose up -d
```

### 方法二：使用Docker命令

```bash
docker run -d \
  --name excel-mcp-server \
  -p 8000:8000 \
  -v /Users:/Users \
  -v $HOME:$HOME \
  -v 您的路径:您的路径 \
  -e EXCEL_FILES_PATH=您的路径 \
  excel-mcp-server
```

## 配置与使用

服务启动后，MCP服务器会在 http://localhost:8000/sse 上运行。

### 在Cursor IDE中使用

在Cursor IDE的配置中添加（每位用户可以设置自己的路径）:

```json
{
  "mcpServers": {
    "excel-mcp-server": {
      "url": "http://localhost:8000/sse",
      "env": {
        "EXCEL_FILES_PATH": "/您选择的路径"
      }
    }
  }
}
```

**重要提示：** 
- macOS用户应该选择`/Users`开头的路径或其主目录下的路径
- 确保选择的路径已在Docker Desktop的文件共享设置中添加
- 修改mcp.json中的路径后，需要重新运行启动脚本让更改生效

## 疑难解答

1. 如果端口8000已被占用，可以在`docker-compose.yml`中修改端口映射:

```yaml
ports:
  - "8001:8000"  # 将主机的8001端口映射到容器的8000端口
```

2. 如果遇到权限问题，可能需要调整您的Excel文件目录的权限:

```bash
chmod 755 /您的路径
```

3. 如果Docker容器无法访问挂载的目录，这是最常见的问题，请确保：
   - 目录已在Docker Desktop的文件共享设置中添加
   - 在macOS上，只能访问`/Users`、`/Volumes`、`/private`、`/tmp`等有限的目录
   - 检查Docker Desktop的Settings > Resources > File sharing部分

4. 错误"Mounts denied"通常表示Docker没有权限访问指定的路径，解决方法：
   - 打开Docker Desktop设置
   - 选择"Resources" > "File sharing"
   - 添加需要访问的目录
   - 应用更改并重启Docker

5. 如果启动脚本无法正确读取mcp.json，请确保:
   - mcp.json位于`~/.cursor/mcp.json`
   - 文件格式正确
   - 考虑安装jq工具(`brew install jq`)以获得更好的JSON解析支持