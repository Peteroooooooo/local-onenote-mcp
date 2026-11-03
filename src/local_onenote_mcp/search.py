"""Full-text search over a local text index that a background thread keeps current."""

from __future__ import annotations

import threading
import time
from datetime import datetime, timezone
from typing import Any

from .hierarchy import Item, of_type, parse_hierarchy, without_recycle_bin
from .onenote import OneNote
from .reader import render_page
from .text_index import TextIndex

STALE_PAGES_READ_INLINE = 60


def page_text(onenote: OneNote, index: TextIndex, page: Item) -> str:
    """Plain body text (with image OCR) of a page, cached until OneNote reports a change."""

    stamp = page.get("lastModifiedTime", "")
    cached = index.get(page["id"], stamp)
    if cached is not None:
        return cached
    text = render_page(onenote.page_xml(page["id"]), "text", include_title=False)
    index.put(page["id"], stamp, text)
    return text


class Indexer:
    """Fills the text index one page at a time in the background.

    It starts only after a tool has talked to OneNote, so starting the MCP server
    never starts OneNote. Each page read holds the bridge for one call, so a tool
    call waits for at most one page.
    """

    def __init__(self, onenote: OneNote, index: TextIndex, *, enabled: bool, refresh_seconds: float) -> None:
        self.onenote = onenote
        self.index = index
        self.enabled = enabled
        self.refresh_seconds = refresh_seconds
        self._lock = threading.Lock()
        self._thread: threading.Thread | None = None
        self._last_start = float("-inf")
        self.indexed = 0
        self.total = 0
        self.last_error: str | None = None

    def kick(self, force: bool = False) -> None:
        if not self.enabled:
            return
        with self._lock:
            if self.running:
                return
            if not force and time.monotonic() - self._last_start < self.refresh_seconds:
                return
            self._last_start = time.monotonic()
            self._thread = threading.Thread(target=self._run, name="onenote-text-index", daemon=True)
            self._thread.start()

    @property
    def running(self) -> bool:
        return self._thread is not None and self._thread.is_alive()

    def _run(self) -> None:
        try:
            everything = parse_hierarchy(self.onenote.hierarchy_xml("", "pages"))
            pages = without_recycle_bin(of_type(everything, "page"))
            self.index.prune({page["id"] for page in of_type(everything, "page")})
            self.total = len(pages)
            self.indexed = 0
            for page in pages:
                if not self.index.is_fresh(page["id"], page.get("lastModifiedTime", "")):
                    try:
                        page_text(self.onenote, self.index, page)
                    except Exception as exc:
                        self.last_error = f"{page.get('path')}: {exc}"
                    time.sleep(0.005)
                self.indexed += 1
        except Exception as exc:
            self.last_error = str(exc)

    def status(self) -> dict[str, Any]:
        return {
            "enabled": self.enabled,
            "running": self.running,
            "indexed_pages": self.indexed,
            "total_pages": self.total,
            "stored_pages": len(self.index),
            "path": str(self.index.path) if self.index.path else None,
            "last_error": self.last_error,
        }


def query_terms(query: str) -> list[str]:
    return [term.casefold() for term in query.split() if term.strip()]


def score(terms: list[str], name: str, path: str, text: str) -> int:
    """0 if any term is missing; otherwise title hits weigh most, then path, then body occurrences."""

    name_l, path_l, text_l = name.casefold(), path.casefold(), text.casefold()
    total = 0
    for term in terms:
        in_name, in_path, body = term in name_l, term in path_l, text_l.count(term)
        if not (in_name or in_path or body):
            return 0
        total += 20 * in_name + 5 * in_path + min(body, 10)
    return total


def snippet(text: str, terms: list[str]) -> str | None:
    lowered = text.casefold()
    for term in terms:
        at = lowered.find(term)
        if at >= 0:
            return text[max(0, at - 160) : at + len(term) + 240].strip()
    return None


def parse_when(value: str) -> datetime | None:
    """ISO date or datetime -> aware UTC datetime. A bare date means local midnight."""

    value = (value or "").strip()
    if not value:
        return None
    try:
        moment = datetime.fromisoformat(value.replace("Z", "+00:00"))
    except ValueError as exc:
        raise ValueError(f"'{value}' is not an ISO date such as 2026-10-01 or 2026-10-01T09:00.") from exc
    if moment.tzinfo is None:
        moment = moment.astimezone()
    return moment.astimezone(timezone.utc)


def modified_at(page: Item) -> datetime | None:
    stamp = page.get("lastModifiedTime")
    if not stamp:
        return None
    try:
        return datetime.fromisoformat(stamp.replace("Z", "+00:00"))
    except ValueError:
        return None


def in_window(page: Item, after: datetime | None, before: datetime | None) -> bool:
    if after is None and before is None:
        return True
    moment = modified_at(page)
    if moment is None:
        return False
    return (after is None or moment >= after) and (before is None or moment < before)


def scan(
    onenote: OneNote,
    index: TextIndex,
    pages: list[Item],
    terms: list[str],
    *,
    limit: int,
    budget_seconds: float,
) -> dict[str, Any]:
    """Rank every page; pages missing from the index are read until the time budget runs out."""

    deadline = time.monotonic() + budget_seconds
    ranked: list[tuple[int, int, Item]] = []
    scanned = 0
    for order, page in enumerate(pages):
        text = index.get(page["id"], page.get("lastModifiedTime", ""))
        if text is None:
            if time.monotonic() > deadline:
                continue
            try:
                text = page_text(onenote, index, page)
            except Exception as exc:
                page["scan_error"] = str(exc)
                text = ""
        scanned += 1
        value = score(terms, page.get("name", ""), page.get("path", ""), text)
        if value:
            page["score"] = value
            ranked.append((-value, order, page))
    ranked.sort(key=lambda entry: entry[:2])
    return {
        "pages": [page for _, _, page in ranked[: max(1, limit)]],
        "total_matches": len(ranked),
        "scanned": scanned,
        "total_pages": len(pages),
        "scan_complete": scanned == len(pages),
    }


def coverage(index: TextIndex, pages: list[Item]) -> int:
    """How many pages are missing or stale in the index."""

    return sum(1 for page in pages if not index.is_fresh(page["id"], page.get("lastModifiedTime", "")))
