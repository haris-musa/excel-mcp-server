#!/bin/bash

# Excel MCP Server 启动脚本

set -e

# 颜色定义
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
BLUE='\033[0;34m'
NC='\033[0m' # No Color

# 日志函数
log_info() {
    echo -e "${GREEN}[INFO]${NC} $1"
}

log_warn() {
    echo -e "${YELLOW}[WARN]${NC} $1"
}

log_error() {
    echo -e "${RED}[ERROR]${NC} $1"
}

# 检查Docker是否安装
check_docker() {
    if ! command -v docker &> /dev/null; then
        log_error "Docker 未安装，请先安装Docker"
        exit 1
    fi
    
    if ! command -v docker-compose &> /dev/null; then
        log_error "Docker Compose 未安装，请先安装Docker Compose"
        exit 1
    fi
}

# 创建必要的目录
create_directories() {
    log_info "创建必要的目录..."
    
    # 从.env文件读取路径配置
    if [ -f .env ]; then
        source .env
    fi
    
    # 使用默认值如果未配置
    HOST_EXCEL_PATH=${HOST_EXCEL_PATH:-./excel_files}
    HOST_LOG_PATH=${HOST_LOG_PATH:-./logs}
    
    mkdir -p "$HOST_EXCEL_PATH"
    mkdir -p "$HOST_LOG_PATH"
    
    log_info "Excel文件目录: $HOST_EXCEL_PATH"
    log_info "日志目录: $HOST_LOG_PATH"
}

# 检查环境变量文件
check_env_file() {
    if [ ! -f .env ]; then
        log_warn ".env 文件不存在，将使用默认配置"
        log_info "您可以复制 env.example 文件来创建 .env 文件"
        
        # 创建默认的.env文件
        cat > .env << EOF
# Excel MCP Server 配置
HOST_EXCEL_PATH=./excel_files
HOST_LOG_PATH=./logs
HOST_PORT=8000
FASTMCP_PORT=8000
EOF
        log_info "已创建默认的 .env 文件"
    fi
}

# 构建并启动服务
start_service() {
    log_info "构建并启动 Excel MCP Server..."
    
    # 构建镜像
    docker-compose build
    
    # 启动服务
    docker-compose up -d
    
    # 等待服务启动
    log_info "等待服务启动..."
    sleep 5
    
    # 检查服务状态
    if docker-compose ps | grep -q "Up"; then
        log_info "服务启动成功！"
        
        # 显示访问信息
        source .env
        HOST_PORT=${HOST_PORT:-8000}
        
        echo ""
        log_info "服务访问信息:"
        log_info "  MCP端点: http://localhost:${HOST_PORT}/mcp"
        log_info "  Excel文件目录: ${HOST_EXCEL_PATH:-./excel_files}"
        log_info "  日志目录: ${HOST_LOG_PATH:-./logs}"
        echo ""
        
        # 显示客户端配置
        log_info "MCP客户端配置:"
        echo -e "${BLUE}{"
        echo -e "   \"mcpServers\": {"
        echo -e "      \"excel\": {"
        echo -e "         \"url\": \"http://localhost:${HOST_PORT}/mcp\""
        echo -e "      }"
        echo -e "   }"
        echo -e "}${NC}"
        
    else
        log_error "服务启动失败！"
        log_error "请检查日志: docker-compose logs excel-mcp-server"
        exit 1
    fi
}

# 显示帮助信息
show_help() {
    echo "Excel MCP Server 启动脚本"
    echo ""
    echo "用法: $0 [选项]"
    echo ""
    echo "选项:"
    echo "  start     启动服务（默认）"
    echo "  stop      停止服务"
    echo "  restart   重启服务"
    echo "  status    查看服务状态"
    echo "  logs      查看服务日志"
    echo "  help      显示此帮助信息"
    echo ""
}

# 主函数
main() {
    case "${1:-start}" in
        start)
            check_docker
            check_env_file
            create_directories
            start_service
            ;;
        stop)
            log_info "停止服务..."
            docker-compose down
            log_info "服务已停止"
            ;;
        restart)
            log_info "重启服务..."
            docker-compose down
            sleep 2
            docker-compose up -d
            log_info "服务已重启"
            ;;
        status)
            log_info "服务状态:"
            docker-compose ps
            ;;
        logs)
            log_info "服务日志:"
            docker-compose logs -f excel-mcp-server
            ;;
        help)
            show_help
            ;;
        *)
            log_error "未知选项: $1"
            show_help
            exit 1
            ;;
    esac
}

# 执行主函数
main "$@" 