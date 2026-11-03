"""Runtime settings, read once from environment variables."""

from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path


def _flag(name: str, default: bool) -> bool:
    value = os.environ.get(name)
    if value is None:
        return default
    return value.strip().lower() not in {"0", "false", "off", "no", ""}


def _number(name: str, default: float) -> float:
    try:
        return float(os.environ.get(name, default))
    except ValueError:
        return default


def data_dir() -> Path:
    base = os.environ.get("LOCALAPPDATA") or str(Path.home() / ".cache")
    return Path(base) / "local-onenote-mcp"


@dataclass(frozen=True)
class Settings:
    timeout_seconds: int
    max_text_chars: int
    hierarchy_cache_seconds: float
    scan_seconds: float
    background_index: bool
    index_refresh_seconds: float
    index_path: Path | None
    backup_path: Path | None
    backup_max_mb: float
    backup_versions: int
    backup_days: float

    @classmethod
    def from_env(cls) -> "Settings":
        root = data_dir()
        index_path = None
        if _flag("LOCAL_ONENOTE_MCP_INDEX", True):
            index_path = Path(os.environ.get("LOCAL_ONENOTE_MCP_INDEX_PATH") or root / "text-index.sqlite3").expanduser()
        backup_path = None
        if _flag("LOCAL_ONENOTE_MCP_BACKUP", True):
            backup_path = Path(os.environ.get("LOCAL_ONENOTE_MCP_BACKUP_PATH") or root / "backups.sqlite3").expanduser()
        return cls(
            timeout_seconds=int(_number("LOCAL_ONENOTE_MCP_TIMEOUT", 90)),
            max_text_chars=int(_number("LOCAL_ONENOTE_MCP_MAX_TEXT_CHARS", 60000)),
            hierarchy_cache_seconds=_number("LOCAL_ONENOTE_MCP_HIERARCHY_CACHE_SECONDS", 5),
            scan_seconds=_number("LOCAL_ONENOTE_MCP_SCAN_SECONDS", 60),
            background_index=_flag("LOCAL_ONENOTE_MCP_BACKGROUND_INDEX", True),
            index_refresh_seconds=_number("LOCAL_ONENOTE_MCP_INDEX_REFRESH_SECONDS", 120),
            index_path=index_path,
            backup_path=backup_path,
            backup_max_mb=_number("LOCAL_ONENOTE_MCP_BACKUP_MB", 200),
            backup_versions=int(_number("LOCAL_ONENOTE_MCP_BACKUP_VERSIONS", 10)),
            backup_days=_number("LOCAL_ONENOTE_MCP_BACKUP_DAYS", 30),
        )
