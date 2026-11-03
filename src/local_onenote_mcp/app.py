"""Shared runtime: settings, the OneNote connection, text index, backups, and the MCP tool registry."""

from __future__ import annotations

import functools
from typing import Any, Callable

import anyio
from mcp.server.fastmcp import FastMCP
from mcp.server.fastmcp.exceptions import ToolError
from mcp.types import ToolAnnotations

from .backup import BackupStore
from .bridge import OneNoteBridge
from .config import Settings
from .onenote import OneNote
from .search import Indexer
from .text_index import TextIndex

INSTRUCTIONS = """\
Controls the local OneNote desktop app on this Windows machine.
- Identifiers accept an object ID, an exact path such as "Notebook/Section/Page", or a unique name.
  IDs are fastest and never ambiguous; page titles are often duplicated, so prefer IDs for pages.
- Read pages with get_page (format="markdown" keeps headings, lists, tables, to-dos and image OCR text).
  To look at an image on a page, pass its callback_id to get_page_image.
- To change part of a page, call get_page(include_paragraphs=true) and then edit_page; this keeps
  everything else on the page (images, handwriting, attachments) untouched.
- search_pages matches every word of the query, ranked; it can also filter by modification date.
- Destructive edits save a backup first; list_backups and restore_backup undo them.
"""

settings = Settings.from_env()
bridge = OneNoteBridge(timeout_seconds=settings.timeout_seconds)
onenote = OneNote(bridge, cache_seconds=settings.hierarchy_cache_seconds)
text_index = TextIndex(settings.index_path)
indexer = Indexer(onenote, text_index, enabled=settings.background_index, refresh_seconds=settings.index_refresh_seconds)
onenote.on_hierarchy_read = indexer.kick
backups = BackupStore(
    settings.backup_path,
    max_bytes=int(settings.backup_max_mb * 2**20),
    versions=settings.backup_versions,
    days=settings.backup_days,
)
mcp = FastMCP("local-onenote", instructions=INSTRUCTIONS)

READ_ONLY = ToolAnnotations(readOnlyHint=True, openWorldHint=False)
WRITES = ToolAnnotations(readOnlyHint=False, destructiveHint=False, openWorldHint=False)
DESTRUCTIVE = ToolAnnotations(readOnlyHint=False, destructiveHint=True, openWorldHint=False)


def ok(**data: Any) -> dict[str, Any]:
    return {"ok": True, **data}


def tool(annotations: ToolAnnotations, *, structured: bool | None = None) -> Callable[[Callable[..., Any]], Callable[..., Any]]:
    """Register a blocking tool body that runs in a worker thread.

    Called directly (as the tests do), the returned coroutine reports failures as
    {"ok": false, "error": ...}. Over MCP the same failure becomes a tool error,
    so the client sees isError=true.
    """

    def decorator(fn: Callable[..., Any]) -> Callable[..., Any]:
        @functools.wraps(fn)
        async def direct(*args: Any, **kwargs: Any) -> Any:
            try:
                return await anyio.to_thread.run_sync(functools.partial(fn, *args, **kwargs))
            except Exception as exc:
                return {"ok": False, "error": str(exc)}

        @functools.wraps(fn)
        async def over_mcp(*args: Any, **kwargs: Any) -> Any:
            result = await direct(*args, **kwargs)
            if isinstance(result, dict) and result.get("ok") is False:
                raise ToolError(result.get("error") or "OneNote operation failed.")
            return result

        mcp.add_tool(over_mcp, annotations=annotations, structured_output=structured)
        return direct

    return decorator
