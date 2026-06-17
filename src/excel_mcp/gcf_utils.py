"""GCF (Graph Compact Format) serialization for MCP tool responses.

When enabled, encodes structured data using GCF's generic profile instead of
JSON, reducing token count by ~60-80% for typical spreadsheet payloads.

Toggle via the EXCEL_MCP_OUTPUT_FORMAT environment variable:
  - "gcf"  (default): use GCF encoding
  - "json": use standard JSON (original behavior)
"""

import json
import os
import logging
from typing import Any

logger = logging.getLogger(__name__)

# Lazy-loaded GCF availability flag
_gcf_available: bool | None = None


def _check_gcf() -> bool:
    """Check whether gcf-python is installed."""
    global _gcf_available
    if _gcf_available is None:
        try:
            from gcf import encode_generic  # noqa: F401
            _gcf_available = True
        except ImportError:
            _gcf_available = False
            logger.info("gcf-python not installed; falling back to JSON output")
    return _gcf_available


def gcf_enabled() -> bool:
    """Return True if GCF output is enabled and available."""
    fmt = os.environ.get("EXCEL_MCP_OUTPUT_FORMAT", "gcf").lower()
    if fmt != "gcf":
        return False
    return _check_gcf()


def serialize(data: Any, *, indent: int = 2) -> str:
    """Serialize *data* as GCF (if enabled) or JSON.

    This is the single serialization entry point that MCP tool handlers should
    call instead of ``json.dumps``.
    """
    if gcf_enabled():
        from gcf import encode_generic
        return encode_generic(data)
    return json.dumps(data, indent=indent, default=str)
