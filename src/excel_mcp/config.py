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


@dataclass(frozen=True)
class Settings:
    allowed_dirs: list[Path] = field(default_factory=list)
    read_only: bool = False
    limits: Limits = field(default_factory=Limits)
    log_level: LogLevel = "WARNING"
