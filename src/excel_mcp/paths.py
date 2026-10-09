"""Resolution and confinement of paths supplied by clients."""

from pathlib import Path

from excel_mcp.errors import PathNotAllowedError

EXCEL_SUFFIXES = frozenset({".xlsx", ".xlsm", ".xltx", ".xltm"})
MACRO_SUFFIXES = frozenset({".xlsm", ".xltm"})
TEMPLATE_SUFFIXES = frozenset({".xltx", ".xltm"})
IMAGE_SUFFIXES = frozenset({".png", ".jpg", ".jpeg"})


class PathPolicy:
    """Turns client-supplied paths into safe absolute paths.

    With allowed directories, every path must resolve (following symlinks)
    inside one of them, and relative paths start from the first one. Without
    allowed directories, only absolute paths are accepted. Files must have an
    Excel extension, so the server can never create or overwrite other files.
    """

    def __init__(self, allowed_dirs: list[Path]) -> None:
        self.allowed_dirs = [directory.resolve() for directory in allowed_dirs]
        self._drives = {directory.drive.casefold() for directory in self.allowed_dirs}

    @property
    def confined(self) -> bool:
        return bool(self.allowed_dirs)

    def resolve(self, raw_path: str) -> Path:
        path = self._resolve(raw_path)
        if path.suffix.lower() not in EXCEL_SUFFIXES:
            allowed = ", ".join(sorted(EXCEL_SUFFIXES))
            raise PathNotAllowedError(f"Path {raw_path!r} must point to an Excel file ({allowed}).")
        return path

    def resolve_image(self, raw_path: str) -> Path:
        path = self._resolve(raw_path)
        if path.suffix.lower() not in IMAGE_SUFFIXES:
            allowed = ", ".join(sorted(IMAGE_SUFFIXES))
            raise PathNotAllowedError(f"Path {raw_path!r} must point to an image ({allowed}).")
        return path

    def resolve_directory(self, raw_path: str) -> Path:
        if not raw_path.strip() and self.confined:
            return self.allowed_dirs[0]
        path = self._resolve(raw_path)
        if not path.is_dir():
            raise PathNotAllowedError(f"Directory {raw_path!r} does not exist.")
        return path

    def display(self, path: Path) -> str:
        """Show a path the way clients should refer to it in later calls."""
        if self.confined and path.is_relative_to(self.allowed_dirs[0]):
            return path.relative_to(self.allowed_dirs[0]).as_posix()
        return str(path)

    def _resolve(self, raw_path: str) -> Path:
        if not raw_path.strip() or "\x00" in raw_path:
            raise PathNotAllowedError("Path must be a non-empty file system path.")
        path = Path(raw_path).expanduser()
        if ":" in str(path)[len(path.drive) :]:
            raise PathNotAllowedError(f"Path {raw_path!r} cannot contain ':' after the drive.")
        if path.drive and not path.root:
            raise PathNotAllowedError(
                f"Path {raw_path!r} is relative to a drive's current directory. Give the full "
                r"path, e.g. 'C:\Reports\q1.xlsx'."
            )
        if not path.is_absolute():
            if not self.confined:
                raise PathNotAllowedError(
                    f"Path {raw_path!r} must be absolute, e.g. 'C:\\Reports\\q1.xlsx' or "
                    "'/home/me/q1.xlsx'."
                )
            path = self.allowed_dirs[0] / path
        outside = f"Path {raw_path!r} is outside the allowed directories."
        # Refuse other drives and network shares before resolving, since resolving
        # a path on a network share connects to that server.
        if self.confined and path.drive.casefold() not in self._drives:
            raise PathNotAllowedError(outside)
        resolved = path.resolve()
        if not self.allows(resolved):
            raise PathNotAllowedError(outside)
        return resolved

    def allows(self, path: Path) -> bool:
        """Whether a path lies in an allowed directory; any path does when unconfined."""
        return not self.confined or any(path.is_relative_to(d) for d in self.allowed_dirs)
