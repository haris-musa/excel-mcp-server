FROM ghcr.io/astral-sh/uv:python3.13-bookworm-slim

WORKDIR /app
ENV UV_COMPILE_BYTECODE=1 UV_LINK_MODE=copy UV_NO_DEV=1

COPY pyproject.toml uv.lock README.md LICENSE ./
COPY src ./src
RUN uv sync --locked

RUN useradd --create-home app && mkdir /data && chown app /data
USER app

ENV EXCEL_FILES_PATH=/data EXCEL_MCP_HOST=0.0.0.0
EXPOSE 8017
CMD ["uv", "run", "--no-sync", "excel-mcp-server", "streamable-http"]
