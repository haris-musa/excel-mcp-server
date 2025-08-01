# 使用官方Python基础镜像
FROM python:3.11-slim

# 设置工作目录
WORKDIR /app

# 设置环境变量
ENV PYTHONUNBUFFERED=1
ENV PYTHONPATH=/app

# 跳过系统依赖安装（使用Python内置模块进行健康检查）

# 复制项目文件
COPY pyproject.toml ./
COPY src/ ./src/
COPY README.md ./
COPY LICENSE ./

# 直接安装项目依赖，使用更宽松的超时设置
RUN pip install --no-cache-dir --timeout=300 --trusted-host pypi.org --trusted-host pypi.python.org --trusted-host files.pythonhosted.org -e .

# 创建Excel文件存储目录
RUN mkdir -p /app/excel_files

# 暴露端口
EXPOSE 8000

# 设置默认环境变量
ENV EXCEL_FILES_PATH=/app/excel_files
ENV FASTMCP_PORT=8000

# 启动命令
CMD ["excel-mcp-server", "streamable-http"] 