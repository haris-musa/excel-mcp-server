"""Size limits on what a workbook zip may unpack to, for every workbook the server opens."""

import zipfile

from excel_mcp.config import Limits
from excel_mcp.errors import LimitExceededError

# Small entries are exempt from the ratio limit: tiny parts compress extremely well.
_RATIO_FLOOR = 1024 * 1024


def check_expansion(infos: list[zipfile.ZipInfo], limits: Limits) -> None:
    """Refuse a zip whose entries, by their headers, unpack too large or too densely."""
    for info in infos:
        check_ratio(info.file_size, info.compress_size, limits)
    if sum(info.file_size for info in infos) > limits.max_unpacked_bytes:
        raise too_large(limits)


def check_ratio(size: int, compressed: int, limits: Limits) -> None:
    if size > max(_RATIO_FLOOR, compressed * limits.max_compression_ratio):
        raise LimitExceededError(
            "The workbook is compressed too densely to be a real workbook: it exceeds the "
            f"safety limit of {limits.max_compression_ratio}:1 on a part."
        )


def too_large(limits: Limits) -> LimitExceededError:
    return LimitExceededError(
        f"The workbook expands to more than {limits.max_unpacked_bytes:,} bytes, over the "
        f"safety limit of {limits.max_unpack_factor} times the file size limit. "
        "--max-file-mb raises it."
    )
