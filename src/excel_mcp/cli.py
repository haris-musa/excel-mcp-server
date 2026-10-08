"""Command line entry point: ``excel-mcp-server stdio`` or ``excel-mcp-server streamable-http``."""

import argparse
import logging
import os
import sys
import warnings
from pathlib import Path
from typing import TextIO

from excel_mcp import __version__
from excel_mcp.config import Limits, Settings
from excel_mcp.server import create_server
from excel_mcp.server.http import is_loopback, serve

DEFAULT_HOST = "127.0.0.1"
DEFAULT_PORT = 8017
DEFAULT_HTTP_DIRECTORY = "./excel_files"

logger = logging.getLogger(__name__)


def main(argv: list[str] | None = None) -> None:
    args = _parser().parse_args(argv)
    settings = _settings(args)
    server = create_server(settings)
    warnings.showwarning = _log_warning

    if args.transport == "stdio":
        if not settings.allowed_dirs:
            print(
                "excel-mcp-server: any Excel file on this computer can be opened. "
                "Pass --allow-dir DIR to limit access to one folder.",
                file=sys.stderr,
            )
        server.run()
        return

    token = os.environ.get("EXCEL_MCP_AUTH_TOKEN") or None
    if not is_loopback(args.host) and token is None and not args.allow_unauthenticated:
        sys.exit(
            f"Refusing to listen on {args.host} without authentication. Set "
            "EXCEL_MCP_AUTH_TOKEN, or pass --allow-unauthenticated if a proxy handles it."
        )
    for directory in settings.allowed_dirs:
        directory.mkdir(parents=True, exist_ok=True)
    serve(server, args.host, args.port, token)


def _log_warning(
    message: Warning | str,
    category: type[Warning],
    filename: str,
    lineno: int,
    file: TextIO | None = None,
    line: str | None = None,
) -> None:
    # Library warnings can quote workbook content (openpyxl echoes defined names),
    # so they are only shown at --log-level DEBUG.
    logger.debug("%s: %s", category.__name__, message)


def _settings(args: argparse.Namespace) -> Settings:
    directories = list(args.allow_dir)
    if env_value := os.environ.get("EXCEL_FILES_PATH"):
        directories += [entry for entry in env_value.split(os.pathsep) if entry.strip()]
    if args.transport == "streamable-http" and not directories:
        directories = [DEFAULT_HTTP_DIRECTORY]
    return Settings(
        allowed_dirs=[Path(directory).expanduser() for directory in directories],
        read_only=args.read_only or os.environ.get("EXCEL_MCP_READ_ONLY") == "1",
        allow_vba_write=args.allow_vba_write or os.environ.get("EXCEL_MCP_ALLOW_VBA_WRITE") == "1",
        limits=Limits(max_file_bytes=args.max_file_mb * 1024 * 1024),
        log_level=args.log_level,
    )


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="excel-mcp-server", description="MCP server for Excel workbooks."
    )
    parser.add_argument("--version", action="version", version=f"%(prog)s {__version__}")

    common = argparse.ArgumentParser(add_help=False)
    common.add_argument(
        "--allow-dir",
        action="append",
        default=[],
        metavar="DIR",
        help="Only allow workbooks inside DIR (repeatable). Relative paths start in the "
        "first DIR. Also read from EXCEL_FILES_PATH.",
    )
    common.add_argument(
        "--read-only",
        action="store_true",
        help="Only register tools that do not change files. Also EXCEL_MCP_READ_ONLY=1.",
    )
    common.add_argument(
        "--allow-vba-write",
        action="store_true",
        help="Add tools that write VBA macro code into .xlsm files (never run by the server). "
        "Off by default and ignored with --read-only. Also EXCEL_MCP_ALLOW_VBA_WRITE=1.",
    )
    common.add_argument(
        "--max-file-mb", type=int, default=100, help="Largest workbook to open (default 100)."
    )
    common.add_argument(
        "--log-level",
        default="WARNING",
        choices=["DEBUG", "INFO", "WARNING", "ERROR", "CRITICAL"],
        help="Log level for messages on stderr (default WARNING).",
    )

    transports = parser.add_subparsers(dest="transport", required=True)
    transports.add_parser("stdio", parents=[common], help="Serve over stdin/stdout (local).")
    http = transports.add_parser(
        "streamable-http", parents=[common], help="Serve over Streamable HTTP (remote)."
    )
    http.add_argument("--host", default=os.environ.get("EXCEL_MCP_HOST", DEFAULT_HOST))
    http.add_argument(
        "--port", type=int, default=int(os.environ.get("EXCEL_MCP_PORT", DEFAULT_PORT))
    )
    http.add_argument(
        "--allow-unauthenticated",
        action="store_true",
        help="Allow a non-local --host without EXCEL_MCP_AUTH_TOKEN.",
    )
    return parser
