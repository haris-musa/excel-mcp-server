#!/bin/bash

# 脚本用于自动从mcp.json读取EXCEL_FILES_PATH并启动Docker容器

# 查找mcp.json文件
MCP_JSON="$HOME/.cursor/mcp.json"
if [ ! -f "$MCP_JSON" ]; then
    echo "错误: 找不到mcp.json文件，请确保Cursor IDE已正确配置"
    exit 1
fi

echo "找到mcp.json文件: $MCP_JSON"

# 使用jq读取EXCEL_FILES_PATH (需要先安装jq: brew install jq)
if command -v jq >/dev/null 2>&1; then
    echo "使用jq解析mcp.json..."
    EXCEL_PATH=$(jq -r '.mcpServers["excel-mcp-server"].env.EXCEL_FILES_PATH // ""' "$MCP_JSON")
    echo "jq解析结果: '$EXCEL_PATH'"
else
    # 如果没有jq，尝试使用grep和sed提取路径
    echo "jq未安装，使用grep和sed解析..."
    grep -A 10 "excel-mcp-server" "$MCP_JSON"
    EXCEL_PATH=$(grep -A 10 "excel-mcp-server" "$MCP_JSON" | grep "EXCEL_FILES_PATH" | sed -E 's/.*"EXCEL_FILES_PATH": "([^"]+)".*/\1/')
    echo "grep解析结果: '$EXCEL_PATH'"
fi

if [ -z "$EXCEL_PATH" ]; then
    echo "警告: 在mcp.json中未找到EXCEL_FILES_PATH，将使用默认路径"
    # 使用默认路径
    EXCEL_PATH="$HOME/excel-files"
fi

echo "使用Excel文件路径: $EXCEL_PATH"

# 确保目录存在并有正确的权限
mkdir -p "$EXCEL_PATH"
chmod 755 "$EXCEL_PATH"

# 导出为环境变量
export EXCEL_FILES_PATH="$EXCEL_PATH"

# 停止并删除旧容器（如果存在）
docker-compose down >/dev/null 2>&1

# 使用环境变量启动docker-compose
EXCEL_FILES_PATH="$EXCEL_PATH" docker-compose up -d

echo "Excel MCP服务已启动，使用路径: $EXCEL_PATH"
echo "查看日志可使用: docker-compose logs -f"