"""Local cache of page text used for full-text search.

Each entry is keyed by page ID and stores OneNote's lastModifiedTime for the
page, so a stale entry is detected without reading the page again. Entries live
in memory and, unless disabled, in a SQLite file.
"""

from __future__ import annotations

import sqlite3
import threading
from pathlib import Path


class TextIndex:
    def __init__(self, path: Path | None) -> None:
        self.path = path
        self._lock = threading.Lock()
        self._entries: dict[str, tuple[str, str]] = {}
        self._db: sqlite3.Connection | None = None
        self._loaded = False

    def _load(self) -> None:
        if self._loaded:
            return
        self._loaded = True
        if self.path is None:
            return
        try:
            self.path.parent.mkdir(parents=True, exist_ok=True)
            db = sqlite3.connect(self.path, check_same_thread=False)
            db.execute("create table if not exists pages (id text primary key, stamp text not null, text text not null)")
            db.commit()
            self._entries = {row[0]: (row[1], row[2]) for row in db.execute("select id, stamp, text from pages")}
            self._db = db
        except sqlite3.Error:
            self._db = None

    def _write(self, sql: str, rows: list[tuple[str, ...]]) -> None:
        if self._db is None or not rows:
            return
        try:
            self._db.executemany(sql, rows)
            self._db.commit()
        except sqlite3.Error:
            pass

    def get(self, page_id: str, stamp: str) -> str | None:
        if not stamp:
            return None
        with self._lock:
            self._load()
            entry = self._entries.get(page_id)
        return entry[1] if entry and entry[0] == stamp else None

    def is_fresh(self, page_id: str, stamp: str) -> bool:
        return self.get(page_id, stamp) is not None

    def put(self, page_id: str, stamp: str, text: str) -> None:
        if not stamp:
            return
        with self._lock:
            self._load()
            self._entries[page_id] = (stamp, text)
            self._write("insert or replace into pages (id, stamp, text) values (?, ?, ?)", [(page_id, stamp, text)])

    def prune(self, live_ids: set[str]) -> int:
        with self._lock:
            self._load()
            gone = [page_id for page_id in self._entries if page_id not in live_ids]
            for page_id in gone:
                del self._entries[page_id]
            self._write("delete from pages where id = ?", [(page_id,) for page_id in gone])
        return len(gone)

    def clear(self) -> None:
        with self._lock:
            self._load()
            self._entries.clear()
            self._write("delete from pages", [()])

    def __len__(self) -> int:
        with self._lock:
            self._load()
            return len(self._entries)
