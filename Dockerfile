FROM python:3.10-slim

WORKDIR /app

# 安装uv
RUN pip install --no-cache-dir uv

# 复制必要的文件
COPY pyproject.toml ./
COPY README.md ./
COPY src/ ./src/

# 创建默认excel文件存储目录
RUN mkdir -p /app/excel_files

# 安装依赖（使用--system参数在系统环境中安装）
RUN uv pip install --system -e .

# 设置默认端口
ENV FASTMCP_PORT=8000
# 默认的Excel文件路径作为备用，但应该优先使用通过环境变量传入的值
# ENV EXCEL_FILES_PATH=/app/excel_files

# 暴露服务端口
EXPOSE 8000

# 启动服务
CMD ["python", "-m", "excel_mcp.__main__"] 