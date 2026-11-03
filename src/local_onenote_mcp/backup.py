"""Bounded, de-duplicated snapshots of pages taken before destructive edits.

Storage layout (one SQLite file):
- blobs: each image, ink stroke set or attached file, stored once by SHA-256.
  Every snapshot of a page that still shows the same picture points at the
  same blob, so repeated edits cost only the page's compressed text XML.
- snapshots: the page XML with binary payloads swapped for blob references,
  zlib-compressed, plus where the page lived and why it was saved.

Limits: at most `versions` snapshots per page, none older than `days`, and the
whole file under `max_bytes` (oldest snapshots go first). A page whose content
matches its newest snapshot is not saved again.
"""

from __future__ import annotations

import base64
import hashlib
import sqlite3
import threading
import time
import xml.etree.ElementTree as ET
import zlib
from pathlib import Path
from typing import Any, Callable

from .constants import ONE
from .hierarchy import parse_xml

BLOB_ATTR = "mcpBlob"
FILE_TAGS = {ONE + "InsertedFile", ONE + "MediaFile"}
MAX_ATTACHMENT_BYTES = 50 * 2**20

SCHEMA = """
create table if not exists blobs (hash text primary key, data blob not null, size integer not null);
create table if not exists snapshots (
    id integer primary key autoincrement,
    page_id text not null, section_id text, path text, title text,
    reason text, created real not null, digest text not null,
    xml blob not null, size integer not null
);
create table if not exists links (snapshot_id integer not null, hash text not null, primary key (snapshot_id, hash));
create index if not exists snapshots_page on snapshots (page_id, created);
"""


class BackupStore:
    def __init__(self, path: Path | None, *, max_bytes: int, versions: int, days: float) -> None:
        self.path = path
        self.max_bytes = max_bytes
        self.versions = max(1, versions)
        self.days = days
        self._lock = threading.Lock()
        self._db: sqlite3.Connection | None = None

    @property
    def enabled(self) -> bool:
        return self.path is not None

    def _conn(self) -> sqlite3.Connection:
        if self._db is None:
            if self.path is None:
                raise RuntimeError("Backups are disabled (LOCAL_ONENOTE_MCP_BACKUP=0).")
            self.path.parent.mkdir(parents=True, exist_ok=True)
            self._db = sqlite3.connect(self.path, check_same_thread=False)
            self._db.execute("pragma auto_vacuum = incremental")
            self._db.executescript(SCHEMA)
        return self._db

    # -- writing ---------------------------------------------------------

    def save(
        self,
        page_xml: str,
        *,
        page_id: str,
        section_id: str = "",
        path: str = "",
        title: str = "",
        reason: str,
        read_file: Callable[[ET.Element], bytes | None] | None = None,
    ) -> int | None:
        """Store a page (XML fetched with page_info=all). Returns the snapshot ID; an unchanged page reuses the last one."""

        if not self.enabled:
            return None
        root = parse_xml(page_xml)
        blobs: dict[str, bytes] = {}
        for data in root.iter(ONE + "Data"):
            if data.text:
                raw = base64.b64decode(data.text)
                digest = hashlib.sha256(raw).hexdigest()
                blobs[digest] = raw
                data.text = None
                data.set(BLOB_ATTR, digest)
        parents = {child: parent for parent in root.iter() for child in parent}
        for node in list(root.iter()):
            if node.tag in FILE_TAGS:
                raw = read_file(node) if read_file is not None else None
                if raw is not None and len(raw) <= MAX_ATTACHMENT_BYTES:
                    digest = hashlib.sha256(raw).hexdigest()
                    blobs[digest] = raw
                    node.set(BLOB_ATTR, digest)
                node.attrib.pop("pathCache", None)
            elif node.tag == ONE + "CallbackID":
                parents[node].remove(node)
        # Temp-file paths and callback IDs change on every read; without them the XML is stable.
        body = ET.tostring(root, encoding="utf-8")
        digest = hashlib.sha256(body).hexdigest()
        with self._lock:
            db = self._conn()
            latest = db.execute("select id, digest from snapshots where page_id = ? order by id desc limit 1", (page_id,)).fetchone()
            if latest and latest[1] == digest:
                return latest[0]
            packed = zlib.compress(body, 6)
            for blob_hash, raw in blobs.items():
                db.execute("insert or ignore into blobs (hash, data, size) values (?, ?, ?)", (blob_hash, raw, len(raw)))
            cursor = db.execute(
                "insert into snapshots (page_id, section_id, path, title, reason, created, digest, xml, size) "
                "values (?, ?, ?, ?, ?, ?, ?, ?, ?)",
                (page_id, section_id, path, title, reason, time.time(), digest, packed, len(packed)),
            )
            snapshot_id = int(cursor.lastrowid)
            db.executemany("insert or ignore into links (snapshot_id, hash) values (?, ?)", [(snapshot_id, h) for h in blobs])
            self._prune(db, page_id)
            db.commit()
            db.execute("pragma incremental_vacuum")
            return snapshot_id

    def _prune(self, db: sqlite3.Connection, page_id: str) -> None:
        doomed = [row[0] for row in db.execute(
            "select id from snapshots where page_id = ? order by id desc limit -1 offset ?", (page_id, self.versions)
        )]
        if self.days > 0:
            cutoff = time.time() - self.days * 86400
            doomed += [row[0] for row in db.execute("select id from snapshots where created < ?", (cutoff,))]
        self._delete(db, doomed)
        while self._total(db) > self.max_bytes:
            oldest = db.execute("select id from snapshots order by id limit 1").fetchone()
            if oldest is None or oldest[0] == db.execute("select max(id) from snapshots").fetchone()[0]:
                break
            self._delete(db, [oldest[0]])

    @staticmethod
    def _delete(db: sqlite3.Connection, ids: list[int]) -> None:
        if not ids:
            return
        db.executemany("delete from snapshots where id = ?", [(i,) for i in ids])
        db.executemany("delete from links where snapshot_id = ?", [(i,) for i in ids])
        db.execute("delete from blobs where hash not in (select hash from links)")

    @staticmethod
    def _total(db: sqlite3.Connection) -> int:
        snapshots = db.execute("select coalesce(sum(size), 0) from snapshots").fetchone()[0]
        blobs = db.execute("select coalesce(sum(size), 0) from blobs").fetchone()[0]
        return int(snapshots + blobs)

    # -- reading ---------------------------------------------------------

    def list(self, page_id: str = "", limit: int = 20) -> list[dict[str, Any]]:
        if not self.enabled or (self.path is not None and not self.path.exists()):
            return []
        query = "select id, page_id, section_id, path, title, reason, created, size from snapshots"
        params: tuple[Any, ...] = ()
        if page_id:
            query += " where page_id = ?"
            params = (page_id,)
        query += " order by id desc limit ?"
        with self._lock:
            rows = self._conn().execute(query, params + (max(1, limit),)).fetchall()
        keys = ("backup_id", "page_id", "section_id", "path", "title", "reason", "created", "bytes")
        result = []
        for row in rows:
            record = dict(zip(keys, row))
            record["created"] = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(record["created"]))
            result.append(record)
        return result

    def load(self, snapshot_id: int) -> tuple[dict[str, Any], ET.Element, dict[str, bytes]]:
        """The snapshot's metadata, its page tree with images restored, and attached files by blob hash."""

        with self._lock:
            db = self._conn()
            row = db.execute(
                "select page_id, section_id, path, title, xml from snapshots where id = ?", (snapshot_id,)
            ).fetchone()
            if row is None:
                raise ValueError(f"No backup with ID {snapshot_id}. Use list_backups to see what is kept.")
            blobs = dict(db.execute(
                "select b.hash, b.data from blobs b join links l on l.hash = b.hash where l.snapshot_id = ?", (snapshot_id,)
            ).fetchall())
        root = ET.fromstring(zlib.decompress(row[4]))
        files: dict[str, bytes] = {}
        for node in root.iter():
            digest = node.attrib.pop(BLOB_ATTR, None)
            if digest is None:
                continue
            if node.tag == ONE + "Data":
                node.text = base64.b64encode(blobs[digest]).decode("ascii")
            else:
                node.set(BLOB_ATTR, digest)
                files[digest] = blobs[digest]
        meta = {"page_id": row[0], "section_id": row[1], "path": row[2], "title": row[3]}
        return meta, root, files

    def stats(self) -> dict[str, Any]:
        if not self.enabled or (self.path is not None and not self.path.exists()):
            return {"enabled": self.enabled, "snapshots": 0, "bytes": 0, "path": str(self.path) if self.path else None}
        with self._lock:
            db = self._conn()
            count = db.execute("select count(*) from snapshots").fetchone()[0]
            total = self._total(db)
        return {
            "enabled": True,
            "snapshots": count,
            "bytes": total,
            "limit_bytes": self.max_bytes,
            "versions_per_page": self.versions,
            "max_age_days": self.days,
            "path": str(self.path),
        }
