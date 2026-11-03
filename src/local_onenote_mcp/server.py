"""MCP server exposing a pure-local Microsoft OneNote control surface."""

from __future__ import annotations

import base64
import functools
import os
import re
import sys
import threading
import time
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Any, Callable

import anyio
from mcp.server.fastmcp import FastMCP
from mcp.server.fastmcp.exceptions import ToolError
from mcp.server.fastmcp.utilities.types import Image
from mcp.types import ToolAnnotations

from . import __version__
from .bridge import OneNoteBridge, OneNoteBridgeError
from .constants import (
    CREATE_FILE_TYPES,
    FILING_LOCATION_TYPES,
    FILING_LOCATIONS,
    HIERARCHY_SCOPES,
    NEW_PAGE_STYLES,
    ONE_NS,
    PAGE_INFO,
    PUBLISH_FORMATS,
    SPECIAL_LOCATIONS,
    XML_SCHEMA_2013,
)
from .image_utils import proportional_dimensions
from .text_index import TextIndex, default_index_path
from .xml_utils import (
    build_image_page_update_xml,
    build_page_update_xml,
    collect_page_objects,
    content_bottom,
    DELETABLE_PAGE_OBJECT_TYPES,
    filter_items,
    parse_hierarchy,
    parse_xml,
    resolve_item,
    text_from_page_xml,
    title_from_page_xml,
)


MCP_NAME = "local-onenote"
DEFAULT_TIMEOUT = int(os.environ.get("LOCAL_ONENOTE_MCP_TIMEOUT", "90"))
MAX_TEXT_CHARS = int(os.environ.get("LOCAL_ONENOTE_MCP_MAX_TEXT_CHARS", "60000"))
HIERARCHY_CACHE_SECONDS = float(os.environ.get("LOCAL_ONENOTE_MCP_HIERARCHY_CACHE_SECONDS", "5"))
SCAN_BUDGET_SECONDS = float(os.environ.get("LOCAL_ONENOTE_MCP_SCAN_SECONDS", "60"))
BACKGROUND_INDEX = os.environ.get("LOCAL_ONENOTE_MCP_BACKGROUND_INDEX", "1").strip().lower() not in {"0", "false", "off", "no"}
INDEX_REFRESH_SECONDS = float(os.environ.get("LOCAL_ONENOTE_MCP_INDEX_REFRESH_SECONDS", "120"))
STALE_PAGES_READ_INLINE = 60
MAX_IMAGE_BYTES = 3_500_000
DEFAULT_MAX_ITEMS = 500

INSTRUCTIONS = """\
Controls the local OneNote desktop app on this Windows machine.
- Identifiers accept an object ID, an exact path such as "Notebook/Section/Page", or a unique name.
  IDs are fastest and never ambiguous; page titles are often duplicated, so prefer IDs for pages.
- Read pages with get_page (format="markdown" keeps headings, lists, tables, to-dos and image OCR text).
  To look at an image on a page, pass its callback_id to get_page_image.
- search_pages matches every word of the query; results are ranked, titles count most.
- Write content as plain, html, or markdown. Writes and deletes change the user's real notes.
"""

mcp = FastMCP(MCP_NAME, instructions=INSTRUCTIONS)
bridge = OneNoteBridge(timeout_seconds=DEFAULT_TIMEOUT)

READ_ONLY = ToolAnnotations(readOnlyHint=True, openWorldHint=False)
WRITES = ToolAnnotations(readOnlyHint=False, destructiveHint=False, openWorldHint=False)
DESTRUCTIVE = ToolAnnotations(readOnlyHint=False, destructiveHint=True, openWorldHint=False)

WRITE_OPERATIONS = {
    "close_notebook",
    "create_new_page",
    "delete_hierarchy",
    "delete_page_content",
    "merge_sections",
    "open_hierarchy",
    "sync_hierarchy",
    "update_hierarchy",
    "update_page_content",
}

ONENOTE_ID = re.compile(r"^\{[0-9A-Fa-f-]{36}\}\{\d+\}\{[0-9A-Za-z]+\}$")
SECTION_LEVEL_TYPES = {"notebook", "section_group", "section"}

_cache_lock = threading.Lock()
_hierarchy_cache: dict[tuple[str, str], tuple[float, str, list[dict[str, Any]]]] = {}
text_index = TextIndex(default_index_path())


def _error(message: str) -> dict[str, Any]:
    return {"ok": False, "error": message}


def _ok(**data: Any) -> dict[str, Any]:
    return {"ok": True, **data}


def tool(annotations: ToolAnnotations, *, structured: bool | None = None) -> Callable[[Callable[..., Any]], Callable[..., Any]]:
    """Register a blocking tool body that runs in a worker thread.

    Called directly (as the tests do), the returned function reports failures as
    {"ok": false, "error": ...}. Over MCP, the same failure is raised as a tool
    error so the client sees isError=true.
    """

    def decorator(fn: Callable[..., Any]) -> Callable[..., Any]:
        @functools.wraps(fn)
        async def wrapper(*args: Any, **kwargs: Any) -> Any:
            try:
                return await anyio.to_thread.run_sync(functools.partial(fn, *args, **kwargs))
            except Exception as exc:
                return _error(str(exc))

        @functools.wraps(fn)
        async def mcp_entry(*args: Any, **kwargs: Any) -> Any:
            result = await wrapper(*args, **kwargs)
            if isinstance(result, dict) and result.get("ok") is False:
                raise ToolError(result.get("error") or "OneNote operation failed.")
            return result

        mcp.add_tool(mcp_entry, annotations=annotations, structured_output=structured)
        return wrapper

    return decorator


def _enum(name: str, value: str, options: dict[str, int]) -> int:
    key = value.casefold()
    if key not in options:
        allowed = ", ".join(sorted(options))
        raise ValueError(f"{name} must be one of: {allowed}")
    return options[key]


def _invalidate_hierarchy_cache() -> None:
    with _cache_lock:
        _hierarchy_cache.clear()


def _bridge(operation: str, **params: Any) -> dict[str, Any]:
    writes = operation in WRITE_OPERATIONS
    try:
        return bridge.call(operation, **params)
    except OneNoteBridgeError as exc:
        raise RuntimeError(str(exc)) from exc
    finally:
        if writes:
            _invalidate_hierarchy_cache()


def _cached_hierarchy(start_id: str, scope: str, fresh: bool) -> tuple[str, list[dict[str, Any]]]:
    key = (start_id, scope)
    now = time.monotonic()
    if not fresh and HIERARCHY_CACHE_SECONDS > 0:
        with _cache_lock:
            cached = _hierarchy_cache.get(key)
        if cached and now - cached[0] < HIERARCHY_CACHE_SECONDS:
            return cached[1], cached[2]
    xml = _bridge(
        "get_hierarchy",
        start_id=start_id,
        scope=_enum("scope", scope, HIERARCHY_SCOPES),
        schema=XML_SCHEMA_2013,
    )["xml"]
    items = parse_hierarchy(xml)
    with _cache_lock:
        _hierarchy_cache[key] = (time.monotonic(), xml, items)
    indexer.kick()
    return xml, items


def _hierarchy_xml(start_id: str = "", scope: str = "pages") -> str:
    return _cached_hierarchy(start_id, scope, fresh=False)[0]


def _hierarchy_items(start_id: str = "", scope: str = "pages", fresh: bool = False) -> list[dict[str, Any]]:
    return [dict(item) for item in _cached_hierarchy(start_id, scope, fresh)[1]]


def _scope_for(item_type: str | None) -> str:
    return "sections" if item_type in SECTION_LEVEL_TYPES else "pages"


def _lookup_by_id(object_id: str) -> dict[str, Any] | None:
    try:
        xml = _bridge("get_hierarchy", start_id=object_id, scope=HIERARCHY_SCOPES["self"], schema=XML_SCHEMA_2013)["xml"]
    except RuntimeError:
        return None
    items = parse_hierarchy(xml)
    return items[0] if items else None


def _resolve_id(identifier: str, item_type: str | None = None) -> str:
    if not identifier:
        return ""
    if ONENOTE_ID.match(identifier.strip()):
        item = _lookup_by_id(identifier.strip())
        if item:
            if item_type and item.get("type") != item_type:
                raise ValueError(f"'{identifier}' is a {item.get('type')}, expected a {item_type}.")
            return item["id"]
    return _resolve_item(identifier, item_type)["id"]


def _resolve_item(identifier: str, item_type: str | None = None) -> dict[str, Any]:
    items = _hierarchy_items("", _scope_for(item_type))
    return resolve_item(items, identifier, item_type)


def _find_item_by_path(path: str, item_type: str | None = None, fresh: bool = False) -> dict[str, Any] | None:
    target = path.casefold()
    for item in _hierarchy_items("", _scope_for(item_type), fresh=fresh):
        if item_type and item.get("type") != item_type:
            continue
        if item.get("path", "").casefold() == target:
            return item
    return None


def _find_item_by_id(object_id: str, item_type: str | None = None, fresh: bool = False) -> dict[str, Any] | None:
    if not object_id:
        return None
    for item in _hierarchy_items("", _scope_for(item_type), fresh=fresh):
        if item_type and item.get("type") != item_type:
            continue
        if item.get("id") == object_id:
            return item
    return None


def _friendly_child_path(parent_path: str, child_name: str) -> str:
    normalized = child_name.replace("\\", "/").strip("/")
    if normalized.lower().endswith(".one"):
        normalized = normalized[:-4]
    return f"{parent_path}/{normalized}" if normalized else parent_path


def _refresh_created_item(
    *,
    expected_path: str,
    item_type: str,
    fallback_id: str = "",
    retries: int = 8,
    delay_seconds: float = 0.5,
) -> dict[str, Any] | None:
    # Page titles are not unique, so a page is only ever matched by its own ID.
    # Sections and section groups are files/folders, unique by path, and
    # OneNote may hand back a stale ID for them right after creation.
    for attempt in range(retries):
        if item_type != "page":
            item = _find_item_by_path(expected_path, item_type, fresh=True)
            if item:
                return item
        item = _find_item_by_id(fallback_id, item_type, fresh=True)
        if item:
            return item
        if attempt + 1 < retries:
            time.sleep(delay_seconds)
    return None


def _create_type_to_item_type(create_type: str) -> str | None:
    key = create_type.casefold()
    if key == "section":
        return "section"
    if key in {"folder", "section_group"}:
        return "section_group"
    if key == "notebook":
        return "notebook"
    return None


def _safe_leaf_name(name: str) -> str:
    cleaned = re.sub(r'[<>:"/\\|?*\x00-\x1f]', " ", name).strip()
    cleaned = re.sub(r"\s+", " ", cleaned)
    if not cleaned:
        raise ValueError("Name cannot be empty.")
    return cleaned


def _truncate(text: str, max_chars: int) -> str:
    if len(text) <= max_chars:
        return text
    return text[:max_chars] + f"\n\n[truncated: {len(text) - max_chars} chars omitted]"


def _without_recycle_bin(items: list[dict[str, Any]]) -> list[dict[str, Any]]:
    return [item for item in items if not _is_recycle_bin_item(item)]


def _is_recycle_bin_item(item: dict[str, Any]) -> bool:
    if item.get("isInRecycleBin") == "true" or item.get("isRecycleBin") == "true":
        return True
    return "OneNote_RecycleBin" in item.get("path", "").split("/")


def _limit(items: list[dict[str, Any]], max_items: int) -> tuple[list[dict[str, Any]], bool]:
    if max_items <= 0 or len(items) <= max_items:
        return items, False
    return items[:max_items], True


def _page_xml(page_id: str, page_info: str = "basic") -> str:
    return _bridge(
        "get_page_content",
        page_id=page_id,
        page_info=_enum("page_info", page_info, PAGE_INFO),
        schema=XML_SCHEMA_2013,
    )["xml"]


def _read_format(value: str) -> str:
    key = (value or "text").strip().casefold()
    if key in {"plain", "text"}:
        return "text"
    if key in {"markdown", "md"}:
        return "markdown"
    raise ValueError("format must be 'text' or 'markdown'.")


def _page_search_text(page: dict[str, Any]) -> str:
    """Plain page text (with OCR) for searching, cached until OneNote reports a change."""

    stamp = page.get("lastModifiedTime", "")
    cached = text_index.get(page["id"], stamp)
    if cached is not None:
        return cached
    text = text_from_page_xml(_page_xml(page["id"], "basic"))
    text_index.put(page["id"], stamp, text)
    return text


def _live_pages(start_id: str = "", include_recycle_bin: bool = False, fresh: bool = False) -> list[dict[str, Any]]:
    pages = filter_items(_hierarchy_items(start_id, "pages", fresh=fresh), "page")
    return pages if include_recycle_bin else _without_recycle_bin(pages)


class BackgroundIndexer:
    """Fills the text index one page at a time while the server is idle.

    It starts only after a tool has talked to OneNote, so launching the MCP
    server never launches OneNote by itself. Each page read takes the bridge
    lock briefly, so tool calls wait for at most one page.
    """

    def __init__(self) -> None:
        self._lock = threading.Lock()
        self._thread: threading.Thread | None = None
        self._last_start = float("-inf")
        self.indexed = 0
        self.total = 0
        self.last_error: str | None = None

    def kick(self, force: bool = False) -> None:
        if not BACKGROUND_INDEX:
            return
        with self._lock:
            if self._thread is not None and self._thread.is_alive():
                return
            if not force and time.monotonic() - self._last_start < INDEX_REFRESH_SECONDS:
                return
            self._last_start = time.monotonic()
            self._thread = threading.Thread(target=self._run, name="onenote-text-index", daemon=True)
            self._thread.start()

    @property
    def running(self) -> bool:
        return self._thread is not None and self._thread.is_alive()

    def _run(self) -> None:
        try:
            pages = _live_pages(fresh=True)
            text_index.prune({page["id"] for page in _live_pages(include_recycle_bin=True)})
            self.total = len(pages)
            self.indexed = sum(1 for page in pages if text_index.is_fresh(page["id"], page.get("lastModifiedTime", "")))
            for page in pages:
                if text_index.is_fresh(page["id"], page.get("lastModifiedTime", "")):
                    continue
                try:
                    _page_search_text(page)
                except Exception as exc:
                    self.last_error = f"{page.get('path')}: {exc}"
                self.indexed += 1
                time.sleep(0.005)
        except Exception as exc:
            self.last_error = str(exc)

    def status(self) -> dict[str, Any]:
        return {
            "enabled": BACKGROUND_INDEX,
            "running": self.running,
            "indexed_pages": self.indexed,
            "total_pages": self.total,
            "stored_pages": len(text_index),
            "path": str(text_index.path) if text_index.path else None,
            "last_error": self.last_error,
        }


indexer = BackgroundIndexer()


def _query_terms(query: str) -> list[str]:
    terms = [term.casefold() for term in query.split() if term.strip()]
    return terms or [query.casefold()]


def _score(terms: list[str], name: str, path: str, text: str) -> int:
    """0 when any term is missing; otherwise title hits weigh most, then body occurrences."""

    name_l, path_l, text_l = name.casefold(), path.casefold(), text.casefold()
    score = 0
    for term in terms:
        in_name = term in name_l
        in_path = term in path_l
        body = text_l.count(term)
        if not (in_name or in_path or body):
            return 0
        score += 20 * in_name + 5 * in_path + min(body, 10)
    return score


def _snippet(text: str, terms: list[str]) -> str | None:
    lowered = text.casefold()
    for term in terms:
        idx = lowered.find(term)
        if idx >= 0:
            start = max(0, idx - 160)
            end = min(len(text), idx + len(term) + 240)
            return text[start:end].strip()
    return None


def _local_text_search(
    start_id: str,
    query: str,
    max_results: int,
    include_recycle_bin: bool,
) -> dict[str, Any]:
    """Rank every page in scope; pages missing from the text index are read within the time budget."""

    pages = _live_pages(start_id, include_recycle_bin)
    terms = _query_terms(query)
    deadline = time.monotonic() + SCAN_BUDGET_SECONDS
    scored: list[tuple[int, int, dict[str, Any]]] = []
    scanned = 0
    for order, page in enumerate(pages):
        text = text_index.get(page["id"], page.get("lastModifiedTime", ""))
        if text is None:
            if time.monotonic() > deadline:
                continue
            try:
                text = _page_search_text(page)
            except Exception as exc:
                page["scan_error"] = str(exc)
                text = ""
        scanned += 1
        score = _score(terms, page.get("name", ""), page.get("path", ""), text)
        if score:
            page["score"] = score
            scored.append((-score, order, page))
    scored.sort(key=lambda entry: entry[:2])
    matches = [page for _, _, page in scored[: max(1, max_results)]]
    return {
        "pages": matches,
        "total_matches": len(scored),
        "scanned": scanned,
        "total_pages": len(pages),
        "scan_complete": scanned == len(pages),
    }


def _index_coverage(start_id: str, include_recycle_bin: bool) -> tuple[int, int]:
    pages = _live_pages(start_id, include_recycle_bin)
    fresh = sum(1 for page in pages if text_index.is_fresh(page["id"], page.get("lastModifiedTime", "")))
    return fresh, len(pages)


def _index_search(start_id: str, query: str) -> list[dict[str, Any]]:
    xml = _bridge(
        "find_pages",
        start_id=start_id,
        query=query,
        include_unindexed=False,
        display=False,
        schema=XML_SCHEMA_2013,
    )["xml"]
    return filter_items(parse_hierarchy(xml), "page")


REPLACE_BODY_OBJECT_TYPES = {"Outline", "Image", "InkDrawing", "FileAttachment", "InsertedFile", "MediaFile"}
IDENTIFIER_RESOLUTION_ORDER = ["id", "exact_path", "unique_name"]
IDENTIFIER_TYPES = {"notebook", "section_group", "section", "page"}
SEARCH_BACKENDS = {"auto", "index", "local_scan"}


@tool(READ_ONLY)
def health_check() -> dict[str, Any]:
    """Verify local OneNote COM access and return a small hierarchy summary."""

    items = _without_recycle_bin(_hierarchy_items("", "sections"))
    notebooks = filter_items(items, "notebook")
    sections = filter_items(items, "section")
    return _ok(
        server=MCP_NAME,
        version=__version__,
        transport="stdio",
        python_executable=sys.executable,
        module_path=str(Path(__file__).resolve()),
        process_cwd=str(Path.cwd()),
        timeout_seconds=DEFAULT_TIMEOUT,
        max_text_chars=MAX_TEXT_CHARS,
        identifier_resolution_order=IDENTIFIER_RESOLUTION_ORDER,
        search_default_backend="auto",
        content_formats=["plain", "html", "markdown"],
        read_formats=["text", "markdown"],
        notebooks=len(notebooks),
        sections=len(sections),
        text_index=indexer.status(),
        write_backend="OneNote desktop COM API (persistent PowerShell bridge)",
    )


@tool(READ_ONLY)
def resolve_identifier(identifier: str, item_type: str = "") -> dict[str, Any]:
    """Resolve a OneNote identifier to one live object before using it in another tool."""

    if not identifier:
        raise ValueError("identifier is required.")
    normalized_type = item_type.strip().casefold() or None
    if normalized_type and normalized_type not in IDENTIFIER_TYPES:
        allowed = ", ".join(sorted(IDENTIFIER_TYPES))
        raise ValueError(f"item_type must be empty or one of: {allowed}")
    item = _resolve_item(identifier, normalized_type)
    return _ok(item=item, identifier_resolution_order=IDENTIFIER_RESOLUTION_ORDER)


@tool(READ_ONLY)
def get_special_locations() -> dict[str, Any]:
    """Return OneNote's local special folders: backup, unfiled, and default notebook folder."""

    locations = {}
    for name, value in SPECIAL_LOCATIONS.items():
        locations[name] = _bridge("get_special_location", location=value)["path"]
    return _ok(locations=locations)


@tool(READ_ONLY)
def list_hierarchy(
    start_identifier: str = "",
    scope: str = "pages",
    include_xml: bool = False,
    include_recycle_bin: bool = False,
    max_items: int = DEFAULT_MAX_ITEMS,
) -> dict[str, Any]:
    """List live OneNote hierarchy objects. Identifiers may be an ID, exact path, or unique name.

    max_items caps the returned list (0 = no cap); use scope="sections" or a start_identifier for large note collections.
    """

    start_id = _resolve_id(start_identifier) if start_identifier else ""
    xml = _hierarchy_xml(start_id, scope)
    items = parse_hierarchy(xml)
    if not include_recycle_bin:
        items = _without_recycle_bin(items)
    total = len(items)
    items, truncated = _limit(items, max_items)
    data = _ok(items=items, count=len(items), total=total, truncated=truncated)
    if include_xml:
        data["xml"] = xml
    return data


@tool(READ_ONLY)
def list_notebooks() -> dict[str, Any]:
    """List live notebooks currently known to the local OneNote desktop app."""

    items = _hierarchy_items("", "notebooks")
    notebooks = filter_items(items, "notebook")
    return _ok(notebooks=notebooks, count=len(notebooks))


@tool(READ_ONLY)
def list_sections(notebook_identifier: str = "", include_recycle_bin: bool = False) -> dict[str, Any]:
    """List sections, optionally restricted to a notebook ID, exact path, or unique name."""

    items = _hierarchy_items("", "sections")
    if not include_recycle_bin:
        items = _without_recycle_bin(items)
    sections = filter_items(items, "section")
    if notebook_identifier:
        notebook = resolve_item(items, notebook_identifier, "notebook")
        prefix = notebook["path"] + "/"
        sections = [section for section in sections if section["path"].startswith(prefix)]
    return _ok(sections=sections, count=len(sections))


@tool(READ_ONLY)
def list_pages(section_identifier: str, include_xml: bool = False, include_recycle_bin: bool = False) -> dict[str, Any]:
    """List pages in a section selected by section ID, exact path, or unique name."""

    section = _resolve_item(section_identifier, "section")
    xml = _hierarchy_xml(section["id"], "pages")
    pages = filter_items(parse_hierarchy(xml), "page")
    if not include_recycle_bin:
        pages = _without_recycle_bin(pages)
    for page in pages:
        if not page.get("path", "").startswith(section["path"] + "/"):
            page["path"] = f"{section['path']}/{page.get('name', '(untitled)')}"
            page["notebook_name"] = section.get("notebook_name")
            page["section_name"] = section.get("section_name")
    data = _ok(section=section, pages=pages, count=len(pages))
    if include_xml:
        data["xml"] = xml
    return data


@tool(READ_ONLY)
def get_page(
    page_identifier: str,
    format: str = "text",
    include_ocr: bool = True,
    include_objects: bool = False,
    include_xml: bool = False,
    page_info: str = "basic",
    max_chars: int = MAX_TEXT_CHARS,
) -> dict[str, Any]:
    """Read a page's title and body.

    format="text" returns plain text; format="markdown" keeps headings, bold/italic, links, lists, to-dos and tables.
    include_ocr adds the text OneNote recognised inside images.
    include_objects lists content objects (outlines, images, ink, attachments) with object_id for delete_page_content
    and callback_id for get_page_image / get_binary_content. include_xml adds the raw OneNote page XML.
    """

    fmt = _read_format(format)
    page = _resolve_item(page_identifier, "page")
    xml = _page_xml(page["id"], page_info)
    text = text_from_page_xml(xml, fmt, include_ocr=include_ocr)
    data = _ok(
        page=page,
        title=title_from_page_xml(xml),
        format=fmt,
        text=_truncate(text, max_chars),
        chars=len(text),
    )
    if include_objects:
        data["objects"] = collect_page_objects(xml)
    if include_xml:
        data["xml"] = xml
    return data


def _image_format(data: bytes) -> str | None:
    if data.startswith(b"\x89PNG\r\n\x1a\n"):
        return "png"
    if data.startswith(b"\xff\xd8"):
        return "jpeg"
    if data.startswith((b"GIF87a", b"GIF89a")):
        return "gif"
    if data[:4] == b"RIFF" and data[8:12] == b"WEBP":
        return "webp"
    return None


@tool(READ_ONLY, structured=False)
def get_page_image(page_identifier: str, callback_id: str) -> Any:
    """Show an image from a page to the model. Take callback_id from get_page(include_objects=true)."""

    page_id = _resolve_id(page_identifier, "page")
    raw = base64.b64decode(_bridge("get_binary_page_content", page_id=page_id, callback_id=callback_id)["base64"])
    fmt = _image_format(raw)
    if fmt is None:
        raise ValueError("This object is not a PNG, JPEG, GIF or WebP image; use get_binary_content to download it.")
    if len(raw) > MAX_IMAGE_BYTES:
        raise ValueError(
            f"Image is {len(raw) // 1024} KB, above the {MAX_IMAGE_BYTES // 1024} KB limit for inline images; "
            "use get_binary_content to download it."
        )
    return Image(data=raw, format=fmt)


@tool(READ_ONLY)
def get_binary_content(page_identifier: str, callback_id: str) -> dict[str, Any]:
    """Download binary page content (image or attached file) as base64 by a callback_id from get_page."""

    page_id = _resolve_id(page_identifier, "page")
    result = _bridge("get_binary_page_content", page_id=page_id, callback_id=callback_id)
    return _ok(base64=result["base64"])


@tool(READ_ONLY)
def search_pages(
    query: str,
    start_identifier: str = "",
    max_results: int = 20,
    include_snippets: bool = True,
    include_unindexed: bool = False,
    include_recycle_bin: bool = False,
    backend: str = "auto",
) -> dict[str, Any]:
    """Search pages for all words of the query, ranked (title hits count most).

    backend="auto" (default) answers from the local text index (page text plus image OCR); pages edited since they were
    indexed are re-read on the spot. While the index is still being built, it asks OneNote's search index and falls back
    to a time-limited text scan. backend="local_scan" (or include_unindexed=true) always scans; backend="index" only asks OneNote.
    """

    if not query.strip():
        raise ValueError("query is required.")
    mode = "local_scan" if include_unindexed else backend.strip().casefold()
    if mode not in SEARCH_BACKENDS:
        raise ValueError(f"backend must be one of: {', '.join(sorted(SEARCH_BACKENDS))}")
    start_id = _resolve_id(start_identifier) if start_identifier else ""
    limit = max(1, max_results)
    scan: dict[str, Any] | None = None
    used = mode
    pages: list[dict[str, Any]] = []
    index_error = None
    use_text_index = False

    if mode == "auto":
        fresh, total = _index_coverage(start_id, include_recycle_bin)
        if total and total - fresh <= STALE_PAGES_READ_INLINE:
            mode = "local_scan"
            use_text_index = True
        else:
            indexer.kick(force=True)
    if mode in {"auto", "index"}:
        try:
            pages = _index_search(start_id, query)
            used = "onenote_index"
        except Exception as exc:
            if mode == "index":
                raise
            index_error = str(exc)
        if not include_recycle_bin:
            pages = _without_recycle_bin(pages)
    if mode == "local_scan" or (mode == "auto" and not pages):
        scan = _local_text_search(start_id, query, limit, include_recycle_bin)
        pages = scan["pages"]
        used = "text_index" if use_text_index else "local_scan"

    pages = pages[:limit]
    if include_snippets:
        terms = _query_terms(query)
        for page in pages:
            try:
                snippet = _snippet(_page_search_text(page), terms)
                if snippet:
                    page["snippet"] = snippet
            except Exception as exc:
                page["snippet_error"] = str(exc)

    data = _ok(pages=pages, count=len(pages), search_backend=used)
    if scan is not None:
        data.update({key: scan[key] for key in ("total_matches", "scanned", "total_pages", "scan_complete")})
    if used != "text_index":
        data["text_index"] = indexer.status()
    if index_error:
        data["index_error"] = index_error
    return data


@tool(READ_ONLY)
def find_meta(start_identifier: str, name: str, include_unindexed: bool = True, include_xml: bool = False) -> dict[str, Any]:
    """Find pages or objects with matching OneNote meta name."""

    start_id = _resolve_id(start_identifier) if start_identifier else ""
    xml = _bridge(
        "find_meta",
        start_id=start_id,
        name=name,
        include_unindexed=include_unindexed,
        schema=XML_SCHEMA_2013,
    )["xml"]
    items = parse_hierarchy(xml)
    data = _ok(items=items, count=len(items))
    if include_xml:
        data["xml"] = xml
    return data


@tool(READ_ONLY)
def get_hyperlink(object_identifier: str, page_content_object_id: str = "", web: bool = False) -> dict[str, Any]:
    """Return a OneNote client or web hyperlink for an object."""

    object_id = _resolve_id(object_identifier)
    operation = "get_web_hyperlink" if web else "get_hyperlink"
    result = _bridge(operation, object_id=object_id, page_content_object_id=page_content_object_id)
    return _ok(hyperlink=result["hyperlink"])


@tool(WRITES)
def open_hierarchy(path: str, relative_to_identifier: str = "", create_type: str = "none") -> dict[str, Any]:
    """Open or create a notebook, section group, or section. Existing OneNote hierarchy paths resolve directly."""

    normalized_create_type = create_type.strip().casefold() or "none"
    relative_to_id = ""
    expected_path = path.replace("\\", "/").strip("/")
    if relative_to_identifier:
        parent = _resolve_item(relative_to_identifier)
        relative_to_id = parent["id"]
        expected_path = _friendly_child_path(parent["path"], path)

    if normalized_create_type == "none":
        existing = _find_item_by_path(expected_path)
        if existing:
            return _ok(object_id=existing["id"], item=existing, opened_existing=True)
        if not relative_to_identifier:
            try:
                existing = _resolve_item(path)
                return _ok(object_id=existing["id"], item=existing, opened_existing=True)
            except Exception:
                pass

    result = _bridge(
        "open_hierarchy",
        path=path,
        relative_to_id=relative_to_id,
        create_file_type=_enum("create_type", normalized_create_type, CREATE_FILE_TYPES),
    )
    item_type = _create_type_to_item_type(normalized_create_type)
    item = (
        _refresh_created_item(expected_path=expected_path, item_type=item_type, fallback_id=result["object_id"])
        if item_type
        else None
    )
    data = _ok(object_id=item["id"] if item else result["object_id"], opened_existing=False)
    if item:
        data["item"] = item
    # OneNote keeps loading a just-opened notebook's sections in the background.
    _invalidate_hierarchy_cache()
    return data


@tool(WRITES)
def create_notebook(name_or_path: str, base_folder: str = "") -> dict[str, Any]:
    """Create a local notebook folder and open it in OneNote."""

    raw = Path(name_or_path)
    if raw.is_absolute():
        notebook_path = raw
    else:
        if base_folder:
            root = Path(base_folder)
        else:
            root = Path(_bridge("get_special_location", location=SPECIAL_LOCATIONS["default_notebook_folder"])["path"])
        notebook_path = root / _safe_leaf_name(name_or_path)
    result = _bridge(
        "open_hierarchy",
        path=str(notebook_path),
        relative_to_id="",
        create_file_type=CREATE_FILE_TYPES["notebook"],
    )
    _invalidate_hierarchy_cache()
    return _ok(path=str(notebook_path), notebook_id=result["object_id"])


@tool(WRITES)
def create_section(parent_identifier: str, section_name: str) -> dict[str, Any]:
    """Create a section under a notebook or section group."""

    parent = _resolve_item(parent_identifier)
    if parent["type"] not in {"notebook", "section_group"}:
        raise ValueError("parent_identifier must resolve to a notebook or section_group.")
    filename = _safe_leaf_name(section_name)
    if not filename.lower().endswith(".one"):
        filename += ".one"
    result = _bridge(
        "open_hierarchy",
        path=filename,
        relative_to_id=parent["id"],
        create_file_type=CREATE_FILE_TYPES["section"],
    )
    expected_path = _friendly_child_path(parent["path"], filename)
    section = _refresh_created_item(expected_path=expected_path, item_type="section", fallback_id=result["object_id"])
    return _ok(
        parent=parent,
        section=section,
        section_id=section["id"] if section else result["object_id"],
        name=section_name,
        path=expected_path,
    )


@tool(WRITES)
def create_section_group(parent_identifier: str, group_name: str) -> dict[str, Any]:
    """Create a section group under a notebook or another section group."""

    parent = _resolve_item(parent_identifier)
    if parent["type"] not in {"notebook", "section_group"}:
        raise ValueError("parent_identifier must resolve to a notebook or section_group.")
    result = _bridge(
        "open_hierarchy",
        path=_safe_leaf_name(group_name),
        relative_to_id=parent["id"],
        create_file_type=CREATE_FILE_TYPES["section_group"],
    )
    expected_path = _friendly_child_path(parent["path"], group_name)
    group = _refresh_created_item(expected_path=expected_path, item_type="section_group", fallback_id=result["object_id"])
    return _ok(
        parent=parent,
        section_group=group,
        section_group_id=group["id"] if group else result["object_id"],
        name=group_name,
        path=expected_path,
    )


@tool(WRITES)
def create_page(
    section_identifier: str,
    title: str,
    content: str = "",
    content_format: str = "plain",
    new_page_style: str = "blank_with_title",
) -> dict[str, Any]:
    """Create a page in a local OneNote section. content_format accepts plain, html, or markdown."""

    section = _resolve_item(section_identifier, "section")
    xml_preview = build_page_update_xml("pending", title=title, content=content, content_format=content_format)
    result = _bridge(
        "create_new_page",
        section_id=section["id"],
        new_page_style=_enum("new_page_style", new_page_style, NEW_PAGE_STYLES),
    )
    page_id = result["page_id"]
    xml = xml_preview.replace('ID="pending"', f'ID="{page_id}"', 1)
    _bridge("update_page_content", xml=xml, schema=XML_SCHEMA_2013, force=False)
    expected_path = _friendly_child_path(section["path"], title)
    page = _refresh_created_item(expected_path=expected_path, item_type="page", fallback_id=page_id)
    return _ok(page_id=page_id, page=page, section=section, title=title, path=page["path"] if page else expected_path)


@tool(WRITES)
def update_page_title(page_identifier: str, title: str) -> dict[str, Any]:
    """Update a page title."""

    page_id = _resolve_id(page_identifier, "page")
    xml = build_page_update_xml(page_id, title=title)
    _bridge("update_page_content", xml=xml, schema=XML_SCHEMA_2013, force=False)
    return _ok(page_id=page_id, title=title)


BLOCK_GAP = 18.0


def _below_existing(page_id: str, x: float | None, y: float | None) -> tuple[float | None, float | None]:
    """Keep caller coordinates; otherwise place a new block under the page's current content."""

    if x is not None or y is not None:
        return x, y
    bottom = content_bottom(_page_xml(page_id, "basic"))
    if bottom is None:
        return None, None
    return 36.0, bottom + BLOCK_GAP


@tool(WRITES)
def append_to_page(
    page_identifier: str,
    content: str,
    content_format: str = "plain",
    x: float | None = None,
    y: float | None = None,
) -> dict[str, Any]:
    """Add a new block to a page, below the existing content unless x/y are given. content_format: plain, html, markdown."""

    page_id = _resolve_id(page_identifier, "page")
    x, y = _below_existing(page_id, x, y)
    xml = build_page_update_xml(page_id, content=content, content_format=content_format, x=x, y=y)
    _bridge("update_page_content", xml=xml, schema=XML_SCHEMA_2013, force=False)
    return _ok(page_id=page_id, appended=True)


@tool(WRITES)
def add_image_to_page(
    page_identifier: str,
    image_path: str,
    image_format: str = "",
    x: float | None = None,
    y: float | None = None,
    width: float | None = None,
    height: float | None = None,
) -> dict[str, Any]:
    """Add a local image file to a page, below the existing content unless x/y are given.

    Give width or height alone to scale proportionally.
    """

    path = Path(image_path)
    if not path.is_file():
        raise ValueError(f"Image file not found: {image_path}")
    fmt = image_format or path.suffix.lstrip(".")
    if not fmt:
        raise ValueError("image_format is required when image_path has no extension.")
    resolved_width, resolved_height = proportional_dimensions(path, width, height)
    image_base64 = base64.b64encode(path.read_bytes()).decode("ascii")
    page_id = _resolve_id(page_identifier, "page")
    x, y = _below_existing(page_id, x, y)
    xml = build_image_page_update_xml(
        page_id,
        image_base64=image_base64,
        image_format=fmt,
        x=x,
        y=y,
        width=resolved_width,
        height=resolved_height,
    )
    _bridge("update_page_content", xml=xml, schema=XML_SCHEMA_2013, force=False)
    return _ok(page_id=page_id, image_path=str(path), width=resolved_width, height=resolved_height)


@tool(DESTRUCTIVE)
def replace_page_body(
    page_identifier: str,
    content: str,
    title: str | None = None,
    content_format: str = "plain",
) -> dict[str, Any]:
    """Replace a page's body content (outlines, images, ink, attachments). content_format accepts plain, html, or markdown.

    The new content is written first and the old objects are deleted afterwards, so a failure never leaves the page empty.
    """

    page_id = _resolve_id(page_identifier, "page")
    old_objects = [
        obj["object_id"]
        for obj in collect_page_objects(_page_xml(page_id, "basic"))
        if obj.get("type") in REPLACE_BODY_OBJECT_TYPES and obj.get("object_id")
    ]
    xml = build_page_update_xml(page_id, title=title, content=content, content_format=content_format)
    _bridge("update_page_content", xml=xml, schema=XML_SCHEMA_2013, force=True)
    deleted = []
    for object_id in old_objects:
        _bridge("delete_page_content", page_id=page_id, object_id=object_id, force=True)
        deleted.append(object_id)
    return _ok(page_id=page_id, deleted_objects=deleted, replaced=True)


@tool(DESTRUCTIVE)
def delete_page_content(page_identifier: str, object_id: str) -> dict[str, Any]:
    """Delete one deletable page content object by object ID. Use get_page(include_objects=true) to find delete_supported objects."""

    page_id = _resolve_id(page_identifier, "page")
    objects = collect_page_objects(_page_xml(page_id, "basic"))
    matched = next((obj for obj in objects if obj.get("object_id") == object_id), None)
    if matched and not matched.get("delete_supported"):
        suggested_id = matched.get("delete_object_id")
        if suggested_id:
            raise ValueError(
                f"Object '{object_id}' is a {matched.get('type')} child and is not directly deletable by OneNote COM. "
                f"Delete its parent content object '{suggested_id}' instead."
            )
        allowed = ", ".join(sorted(DELETABLE_PAGE_OBJECT_TYPES))
        raise ValueError(
            f"Object '{object_id}' is a {matched.get('type')} child and is not directly deletable by OneNote COM. "
            f"Deletable object types: {allowed}."
        )
    _bridge("delete_page_content", page_id=page_id, object_id=object_id, force=True)
    return _ok(page_id=page_id, object_id=object_id, deleted=True)


@tool(DESTRUCTIVE)
def delete_hierarchy(object_identifier: str, permanently: bool = False) -> dict[str, Any]:
    """Delete a notebook, section group, section, or page. permanently=false moves it to OneNote's recycle bin."""

    item = _resolve_item(object_identifier)
    if item["type"] == "page":
        _bridge("delete_hierarchy", object_id=item["id"], permanently=permanently)
        time.sleep(0.5)
        remaining = _find_item_by_id(item["id"], "page")
        if remaining and remaining.get("path") == item["path"]:
            raise RuntimeError(f"Delete returned success, but page '{item['path']}' ({item['id']}) still exists.")
        return _ok(object_id=item["id"], deleted_ids=[item["id"]], permanently=permanently, deleted=True, verified_gone=True)

    # OneNote can re-create a section or section group under a new ID right
    # after deleting it; those names are unique per folder, so retry by path.
    deleted_ids = []
    for attempt in range(4):
        object_id = item["id"]
        _bridge("delete_hierarchy", object_id=object_id, permanently=permanently)
        deleted_ids.append(object_id)
        time.sleep(0.5)
        remaining = _find_item_by_path(item["path"], item["type"])
        if not remaining:
            return _ok(
                object_id=object_id,
                deleted_ids=deleted_ids,
                permanently=permanently,
                deleted=True,
                verified_gone=True,
            )
        item = remaining
        if attempt == 3:
            raise RuntimeError(f"Delete returned success, but '{item['path']}' still exists with ID {item['id']}.")
    raise RuntimeError("Delete did not complete.")


@tool(DESTRUCTIVE)
def update_page_xml(xml: str, force: bool = False) -> dict[str, Any]:
    """Advanced: submit raw OneNote page XML to UpdatePageContent."""

    _bridge("update_page_content", xml=xml, schema=XML_SCHEMA_2013, force=force)
    return _ok(updated=True, force=force)


@tool(DESTRUCTIVE)
def update_hierarchy_xml(xml: str) -> dict[str, Any]:
    """Advanced: submit raw OneNote hierarchy XML to UpdateHierarchy."""

    _bridge("update_hierarchy", xml=xml, schema=XML_SCHEMA_2013)
    return _ok(updated=True)


@tool(WRITES)
def publish_object(object_identifier: str, target_path: str, format: str = "pdf", overwrite: bool = False) -> dict[str, Any]:
    """Export a notebook, section, or page to a local file."""

    output = Path(target_path).expanduser()
    if not output.is_absolute():
        output = Path.cwd() / output
    output = output.resolve(strict=False)
    if output.exists() and not overwrite:
        raise ValueError(f"Target already exists: {target_path}")
    output.parent.mkdir(parents=True, exist_ok=True)
    object_id = _resolve_id(object_identifier)
    result = _bridge(
        "publish",
        object_id=object_id,
        target_path=str(output),
        format=_enum("format", format, PUBLISH_FORMATS),
    )
    return _ok(path=result["path"])


@tool(WRITES)
def navigate_to(object_identifier: str, page_content_object_id: str = "", new_window: bool = False) -> dict[str, Any]:
    """Open a OneNote object in the desktop app."""

    object_id = _resolve_id(object_identifier)
    _bridge("navigate_to", object_id=object_id, page_content_object_id=page_content_object_id, new_window=new_window)
    return _ok(navigated=True)


@tool(WRITES)
def navigate_to_url(url: str, new_window: bool = False) -> dict[str, Any]:
    """Open a OneNote URL in the desktop app."""

    _bridge("navigate_to_url", url=url, new_window=new_window)
    return _ok(navigated=True)


@tool(WRITES)
def sync_hierarchy(object_identifier: str) -> dict[str, Any]:
    """Ask OneNote to sync a notebook hierarchy object."""

    object_id = _resolve_id(object_identifier)
    _bridge("sync_hierarchy", object_id=object_id)
    return _ok(object_id=object_id, synced=True)


@tool(DESTRUCTIVE)
def merge_sections(source_section_identifier: str, destination_section_identifier: str) -> dict[str, Any]:
    """Merge one section into another."""

    source_id = _resolve_id(source_section_identifier, "section")
    destination_id = _resolve_id(destination_section_identifier, "section")
    _bridge("merge_sections", source_section_id=source_id, destination_section_id=destination_id)
    return _ok(source_section_id=source_id, destination_section_id=destination_id, merged=True)


@tool(WRITES)
def close_notebook(notebook_identifier: str, force: bool = False) -> dict[str, Any]:
    """Close a notebook in the desktop OneNote app."""

    notebook_id = _resolve_id(notebook_identifier, "notebook")
    _bridge("close_notebook", notebook_id=notebook_id, force=force)
    return _ok(notebook_id=notebook_id, closed=True)


ONE = f"{{{ONE_NS}}}"
FILE_MOVE_WAIT_SECONDS = 20.0


def _tree(start_id: str, scope: str) -> ET.Element:
    return parse_xml(
        _bridge("get_hierarchy", start_id=start_id, scope=HIERARCHY_SCOPES[scope], schema=XML_SCHEMA_2013)["xml"]
    )


def _submit_tree(tree: ET.Element) -> None:
    _bridge("update_hierarchy", xml=ET.tostring(tree, encoding="unicode"), schema=XML_SCHEMA_2013)


def _page_nodes(section: ET.Element) -> list[ET.Element]:
    return [node for node in section if node.tag == ONE + "Page"]


def _level(node: ET.Element) -> int:
    try:
        return int(node.attrib.get("pageLevel", "1"))
    except ValueError:
        return 1


def _page_block(pages: list[ET.Element], index: int, include_subpages: bool) -> list[ET.Element]:
    """The page at index plus, optionally, the subpages that directly follow it."""

    block = [pages[index]]
    if include_subpages:
        root_level = _level(pages[index])
        for node in pages[index + 1 :]:
            if _level(node) <= root_level:
                break
            block.append(node)
    return block


def _notebook_id_of(item: dict[str, Any]) -> str:
    if item.get("type") == "notebook":
        return item["id"]
    name = item.get("path", "").split("/", 1)[0]
    for notebook in filter_items(_hierarchy_items("", "notebooks", fresh=True), "notebook"):
        if notebook.get("path") == name:
            return notebook["id"]
    raise RuntimeError(f"Could not find the notebook that contains '{item.get('path')}'.")


def _wait_for_file_move(old_path: str) -> bool:
    """Local sections are moved on disk lazily; wait until OneNote has moved the old .one file."""

    if not old_path or "://" in old_path:
        return True
    deadline = time.monotonic() + FILE_MOVE_WAIT_SECONDS
    while time.monotonic() < deadline:
        if not Path(old_path).exists():
            return True
        time.sleep(0.25)
    return not Path(old_path).exists()


def _sync_and_settle(notebook_ids: set[str], old_file: str) -> bool:
    for notebook_id in notebook_ids:
        _bridge("sync_hierarchy", object_id=notebook_id)
    return _wait_for_file_move(old_file)


@tool(WRITES)
def move_page(
    page_identifier: str,
    destination_section_identifier: str = "",
    after_page_identifier: str = "",
    to_top: bool = False,
    include_subpages: bool = True,
) -> dict[str, Any]:
    """Move a page to another section (any notebook) or to a new position in its own section.

    Leave destination_section_identifier empty to reorder within the current section. The page goes after
    after_page_identifier, to the top with to_top=true, or to the end. Its subpages move with it unless
    include_subpages=false. The moved page becomes a top-level page (use set_page_level to indent it).
    Moving to another section gives the page a new ID; the new IDs are returned.
    """

    page = _resolve_item(page_identifier, "page")
    if _is_recycle_bin_item(page):
        raise ValueError("Pages in the recycle bin cannot be moved; restore them in OneNote first.")
    source_id = page["parent_id"]
    destination = _resolve_item(destination_section_identifier, "section") if destination_section_identifier else None
    destination_id = destination["id"] if destination else source_id
    if destination and _is_recycle_bin_item(destination):
        raise ValueError("Use delete_hierarchy to move a page to the recycle bin.")

    source_tree = _tree(source_id, "pages")
    source_pages = _page_nodes(source_tree)
    index = next((i for i, node in enumerate(source_pages) if node.attrib.get("ID") == page["id"]), None)
    if index is None:
        raise RuntimeError(f"Page '{page['path']}' was not found in its section; refresh and try again.")
    block = _page_block(source_pages, index, include_subpages)
    block_ids = [node.attrib.get("ID") for node in block]

    same_section = destination_id == source_id
    target = source_tree if same_section else _tree(destination_id, "pages")
    for node in block:
        if same_section:
            target.remove(node)
    before_ids = {node.attrib.get("ID") for node in _page_nodes(target)}

    shift = _level(block[0]) - 1
    for node in block:
        node.set("pageLevel", str(max(1, _level(node) - shift)))

    target_pages = _page_nodes(target)
    children = list(target)
    if after_page_identifier:
        after_id = _resolve_id(after_page_identifier, "page")
        if after_id in block_ids:
            raise ValueError("after_page_identifier cannot be the page being moved or one of its subpages.")
        after_index = next((i for i, node in enumerate(target_pages) if node.attrib.get("ID") == after_id), None)
        if after_index is None:
            raise ValueError("after_page_identifier must be a page in the destination section.")
        anchor = _page_block(target_pages, after_index, True)[-1]
        insert_at = children.index(anchor) + 1
    elif to_top and target_pages:
        insert_at = children.index(target_pages[0])
    else:
        insert_at = children.index(target_pages[-1]) + 1 if target_pages else len(children)
    for offset, node in enumerate(block):
        target.insert(insert_at + offset, node)

    _submit_tree(target)

    if same_section:
        moved = [{"old_id": page_id, "new_id": page_id} for page_id in block_ids]
    else:
        new_ids: list[str] = []
        for _ in range(20):
            after_nodes = _page_nodes(_tree(destination_id, "pages"))
            new_ids = [node.attrib.get("ID", "") for node in after_nodes if node.attrib.get("ID") not in before_ids]
            if len(new_ids) >= len(block):
                break
            time.sleep(0.25)
        leftovers = {node.attrib.get("ID") for node in _page_nodes(_tree(source_id, "pages"))} & set(block_ids)
        if len(new_ids) < len(block) or leftovers:
            raise RuntimeError(
                f"OneNote accepted the move but the result could not be confirmed "
                f"(found {len(new_ids)} of {len(block)} pages in the destination; {len(leftovers)} still in the source)."
            )
        moved = [{"old_id": old, "new_id": new} for old, new in zip(block_ids, new_ids)]
    _invalidate_hierarchy_cache()
    first = _lookup_by_id(moved[0]["new_id"]) or {"id": moved[0]["new_id"]}
    return _ok(page_id=moved[0]["new_id"], page=first, moved_pages=moved, destination_section_id=destination_id)


@tool(WRITES)
def set_page_level(page_identifier: str, level: int) -> dict[str, Any]:
    """Indent a page as a subpage (level 2 or 3) or promote it back to a top-level page (level 1)."""

    if level not in {1, 2, 3}:
        raise ValueError("level must be 1, 2 or 3.")
    page = _resolve_item(page_identifier, "page")
    tree = _tree(page["parent_id"], "pages")
    pages = _page_nodes(tree)
    index = next((i for i, node in enumerate(pages) if node.attrib.get("ID") == page["id"]), None)
    if index is None:
        raise RuntimeError(f"Page '{page['path']}' was not found in its section; refresh and try again.")
    if level > 1 and (index == 0 or _level(pages[index - 1]) < level - 1):
        raise ValueError("A subpage needs a page directly above it one level higher; move the page first.")
    pages[index].set("pageLevel", str(level))
    _submit_tree(tree)
    return _ok(page_id=page["id"], level=level)


def _container_children(tree: ET.Element, tag: str) -> list[ET.Element]:
    return [
        node
        for node in tree
        if node.tag == ONE + tag and node.attrib.get("isRecycleBin") != "true"
    ]


@tool(WRITES)
def move_section(
    section_identifier: str,
    destination_identifier: str = "",
    after_identifier: str = "",
    to_top: bool = False,
) -> dict[str, Any]:
    """Move a section or section group into a notebook or section group (any notebook), or reorder it in place.

    Leave destination_identifier empty to reorder within the current parent. The item goes after
    after_identifier (a sibling of the same type), to the top with to_top=true, or to the end. IDs are kept.
    For local notebooks the tool waits until OneNote has moved the .one file on disk.
    """

    item = _resolve_item(section_identifier)
    if item["type"] not in {"section", "section_group"}:
        raise ValueError("section_identifier must resolve to a section or section_group.")
    if _is_recycle_bin_item(item):
        raise ValueError("The recycle bin and its contents cannot be moved.")
    source_parent = item["parent_id"]
    destination = _resolve_item(destination_identifier) if destination_identifier else _lookup_by_id(source_parent)
    if destination is None or destination["type"] not in {"notebook", "section_group"}:
        raise ValueError("destination_identifier must resolve to a notebook or section_group.")
    if _is_recycle_bin_item(destination):
        raise ValueError("Use delete_hierarchy to move a section to the recycle bin.")
    if destination["id"] == item["id"] or destination["path"].startswith(item["path"] + "/"):
        raise ValueError("A section group cannot be moved into itself.")

    tag = "Section" if item["type"] == "section" else "SectionGroup"
    same_parent = destination["id"] == source_parent
    target = _tree(destination["id"], "sections")
    if same_parent:
        node = next(child for child in target if child.attrib.get("ID") == item["id"])
        target.remove(node)
    else:
        clash = [child for child in _container_children(target, tag) if child.attrib.get("name", "").casefold() == item["name"].casefold()]
        if clash:
            raise ValueError(f"'{destination['path']}' already contains a {item['type']} named '{item['name']}'; rename one first.")
        node = _tree(item["id"], "sections")

    siblings = _container_children(target, tag)
    children = list(target)
    if after_identifier:
        after_id = _resolve_id(after_identifier)
        anchor = next((child for child in siblings if child.attrib.get("ID") == after_id), None)
        if anchor is None:
            raise ValueError(f"after_identifier must be a {item['type']} directly inside '{destination['path']}'.")
        insert_at = children.index(anchor) + 1
    elif to_top and siblings:
        insert_at = children.index(siblings[0])
    elif siblings:
        insert_at = children.index(siblings[-1]) + 1
    else:
        groups_or_bin = [child for child in children if child.tag == ONE + "SectionGroup"]
        insert_at = children.index(groups_or_bin[0]) if tag == "Section" and groups_or_bin else len(children)
        if tag == "SectionGroup":
            recycle = [child for child in children if child.attrib.get("isRecycleBin") == "true"]
            insert_at = children.index(recycle[0]) if recycle else len(children)
    target.insert(insert_at, node)
    _submit_tree(target)

    file_settled = True
    if not same_parent:
        notebooks = {_notebook_id_of(item), _notebook_id_of(destination)}
        file_settled = _sync_and_settle(notebooks, item.get("onenote_path", ""))
    _invalidate_hierarchy_cache()
    moved = _lookup_by_id(item["id"])
    expected_prefix = destination["path"] + "/"
    refreshed = _find_item_by_id(item["id"], item["type"], fresh=True)
    if refreshed is None or not refreshed["path"].startswith(expected_prefix):
        raise RuntimeError("OneNote accepted the move but the section is not under the destination yet; check OneNote.")
    data = _ok(object_id=item["id"], item=refreshed or moved, destination=destination["path"])
    if not file_settled:
        data["warning"] = (
            "OneNote has not finished moving the section file on disk. Keep OneNote open for a moment before "
            "closing the notebook, or the section can reappear in its old place."
        )
    return data


@tool(WRITES)
def rename_section(section_identifier: str, new_name: str) -> dict[str, Any]:
    """Rename a section or section group. Rename pages with update_page_title."""

    item = _resolve_item(section_identifier)
    if item["type"] not in {"section", "section_group"}:
        raise ValueError("section_identifier must resolve to a section or section_group.")
    if _is_recycle_bin_item(item):
        raise ValueError("Recycle bin items cannot be renamed.")
    name = _safe_leaf_name(new_name)
    tag = "Section" if item["type"] == "section" else "SectionGroup"
    parent = _tree(item["parent_id"], "sections")
    for child in _container_children(parent, tag):
        if child.attrib.get("ID") != item["id"] and child.attrib.get("name", "").casefold() == name.casefold():
            raise ValueError(f"A {item['type']} named '{name}' already exists next to it.")
    node = next(child for child in parent if child.attrib.get("ID") == item["id"])
    node.set("name", name)
    _submit_tree(parent)
    file_settled = _sync_and_settle({_notebook_id_of(item)}, item.get("onenote_path", ""))
    refreshed = _find_item_by_id(item["id"], item["type"], fresh=True)
    data = _ok(object_id=item["id"], item=refreshed, name=name)
    if not file_settled:
        data["warning"] = "OneNote has not finished renaming the section file on disk yet."
    return data


@tool(WRITES)
def set_filing_location(filing_location: str, filing_location_type: str, section_or_page_identifier: str) -> dict[str, Any]:
    """Set OneNote's local filing location for email, web clips, printouts, and similar content."""

    object_id = _resolve_id(section_or_page_identifier)
    _bridge(
        "set_filing_location",
        filing_location=_enum("filing_location", filing_location, FILING_LOCATIONS),
        filing_location_type=_enum("filing_location_type", filing_location_type, FILING_LOCATION_TYPES),
        section_or_page_id=object_id,
    )
    return _ok(object_id=object_id, updated=True)


def main() -> None:
    mcp.run(transport="stdio")


if __name__ == "__main__":
    main()
