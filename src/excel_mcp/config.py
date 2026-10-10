"""Server settings."""

from dataclasses import dataclass, field
from pathlib import Path
from typing import Literal

LogLevel = Literal["DEBUG", "INFO", "WARNING", "ERROR", "CRITICAL"]


@dataclass(frozen=True)
class Limits:
    max_file_bytes: int = 100 * 1024 * 1024
    max_read_cells: int = 10_000
    max_cells: int = 100_000
    max_copy_cells: int = 1_000_000
    # Zip bomb guards, applied to every workbook the server opens or receives. Real workbooks
    # compress 10 to 30 times and sparse sheets far more, so these leave room for them.
    max_unpack_factor: int = 10
    max_compression_ratio: int = 500

    @property
    def max_unpacked_bytes(self) -> int:
        """The most any workbook may expand to, summed over all its parts."""
        return self.max_file_bytes * self.max_unpack_factor


@dataclass(frozen=True)
class Settings:
    allowed_dirs: list[Path] = field(default_factory=list)
    read_only: bool = False
    allow_vba_write: bool = False
    limits: Limits = field(default_factory=Limits)
    log_level: LogLevel = "WARNING"
