"""One object that owns every call to OneNote: bridge, hierarchy cache, identifier resolution, page I/O."""

from __future__ import annotations

import threading
import time
import xml.etree.ElementTree as ET
from typing import Any, Callable

from .bridge import OneNoteBridge, OneNoteBridgeError
from .constants import HIERARCHY_SCOPES, PAGE_INFO, XML_SCHEMA_2013, enum_value
from .hierarchy import ONENOTE_ID, Item, of_type, parse_hierarchy, parse_xml, resolve_item, without_recycle_bin

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
CONTAINER_TYPES = {"notebook", "section_group", "section"}
# OneNote sometimes answers REGDB_E_CLASSNOTREG for a moment while it rewrites notebook files.
# These calls are safe to send twice: reads, and hierarchy updates that describe the wanted end state.
CLASS_NOT_REGISTERED = -2147221164
RETRYABLE_OPERATIONS = {
    "get_binary_page_content",
    "get_hierarchy",
    "get_hyperlink",
    "get_page_content",
    "get_special_location",
    "get_web_hyperlink",
    "find_pages",
    "sync_hierarchy",
    "update_hierarchy",
}


class OneNote:
    def __init__(self, bridge: OneNoteBridge, cache_seconds: float = 5.0) -> None:
        self.bridge = bridge
        self.cache_seconds = cache_seconds
        self._lock = threading.Lock()
        self._cache: dict[tuple[str, str], tuple[float, list[Item]]] = {}
        self.on_hierarchy_read: Callable[[], None] | None = None

    # -- raw calls -------------------------------------------------------

    def call(self, operation: str, **params: Any) -> dict[str, Any]:
        try:
            try:
                return self.bridge.call(operation, **params)
            except OneNoteBridgeError as exc:
                if exc.hresult != CLASS_NOT_REGISTERED or operation not in RETRYABLE_OPERATIONS:
                    raise
                time.sleep(0.5)
                return self.bridge.call(operation, **params)
        except OneNoteBridgeError as exc:
            raise RuntimeError(str(exc)) from exc
        finally:
            if operation in WRITE_OPERATIONS:
                self.invalidate()

    def invalidate(self) -> None:
        with self._lock:
            self._cache.clear()

    # -- hierarchy -------------------------------------------------------

    def hierarchy_xml(self, start_id: str = "", scope: str = "pages") -> str:
        return self.call(
            "get_hierarchy",
            start_id=start_id,
            scope=enum_value("scope", scope, HIERARCHY_SCOPES),
            schema=XML_SCHEMA_2013,
        )["xml"]

    def tree(self, start_id: str, scope: str) -> ET.Element:
        return parse_xml(self.hierarchy_xml(start_id, scope))

    def submit_tree(self, tree: ET.Element) -> None:
        self.call("update_hierarchy", xml=ET.tostring(tree, encoding="unicode"), schema=XML_SCHEMA_2013)

    def items(self, start_id: str = "", scope: str = "pages", fresh: bool = False) -> list[Item]:
        key = (start_id, scope)
        if not fresh and self.cache_seconds > 0:
            with self._lock:
                cached = self._cache.get(key)
            if cached and time.monotonic() - cached[0] < self.cache_seconds:
                return [dict(item) for item in cached[1]]
        items = parse_hierarchy(self.hierarchy_xml(start_id, scope))
        with self._lock:
            self._cache[key] = (time.monotonic(), items)
        if self.on_hierarchy_read:
            self.on_hierarchy_read()
        return [dict(item) for item in items]

    def pages(self, start_id: str = "", include_recycle_bin: bool = False, fresh: bool = False) -> list[Item]:
        pages = of_type(self.items(start_id, "pages", fresh), "page")
        return pages if include_recycle_bin else without_recycle_bin(pages)

    def lookup(self, object_id: str) -> Item | None:
        """One object by ID without scanning the whole hierarchy (a few ms)."""

        try:
            items = parse_hierarchy(self.hierarchy_xml(object_id, "self"))
        except RuntimeError:
            return None
        return items[0] if items else None

    def resolve(self, identifier: str, item_type: str | None = None, full: bool = False) -> Item:
        """ID, exact path, or unique name -> live item.

        For an ID the result comes from a direct lookup (a few ms) and has no path
        or parent; pass full=True when those are needed.
        """

        identifier = (identifier or "").strip()
        if not identifier:
            raise ValueError("An identifier is required.")
        if ONENOTE_ID.match(identifier):
            item = self.lookup(identifier)
            if item:
                if item_type and item["type"] != item_type:
                    raise ValueError(f"'{identifier}' is a {item['type']}, expected a {item_type}.")
                if not full:
                    return item
                return self.find(id=item["id"], item_type=item["type"]) or self.find(id=item["id"], item_type=item["type"], fresh=True) or item
        scope = "sections" if item_type in CONTAINER_TYPES else "pages"
        return resolve_item(self.items("", scope), identifier, item_type)

    def find(self, *, id: str = "", path: str = "", item_type: str | None = None, fresh: bool = False) -> Item | None:
        scope = "sections" if item_type in CONTAINER_TYPES else "pages"
        for item in self.items("", scope, fresh=fresh):
            if item_type and item["type"] != item_type:
                continue
            if (id and item["id"] == id) or (path and item["path"].casefold() == path.casefold()):
                return item
        return None

    def wait_for(self, *, item_type: str, id: str = "", path: str = "", tries: int = 8, delay: float = 0.5) -> Item | None:
        """Re-read until a just-created object shows up. Pages match by ID only (titles repeat)."""

        for attempt in range(tries):
            if path and item_type != "page":
                item = self.find(path=path, item_type=item_type, fresh=True)
                if item:
                    return item
            item = self.find(id=id, item_type=item_type, fresh=True) if id else None
            if item:
                return item
            if attempt + 1 < tries:
                time.sleep(delay)
        return None

    def notebook_of(self, item: Item) -> str:
        if item["type"] == "notebook":
            return item["id"]
        name = item["path"].split("/", 1)[0]
        for notebook in of_type(self.items("", "notebooks", fresh=True), "notebook"):
            if notebook["path"] == name:
                return notebook["id"]
        raise RuntimeError(f"Could not find the notebook that contains '{item['path']}'.")

    # -- pages -----------------------------------------------------------

    def page_xml(self, page_id: str, page_info: str = "basic") -> str:
        return self.call(
            "get_page_content",
            page_id=page_id,
            page_info=enum_value("page_info", page_info, PAGE_INFO),
            schema=XML_SCHEMA_2013,
        )["xml"]

    def update_page(self, xml: str | ET.Element, force: bool = False) -> None:
        if isinstance(xml, ET.Element):
            xml = ET.tostring(xml, encoding="unicode")
        self.call("update_page_content", xml=xml, schema=XML_SCHEMA_2013, force=force)

    def delete_object(self, page_id: str, object_id: str) -> None:
        self.call("delete_page_content", page_id=page_id, object_id=object_id, force=True)

    def new_page(self, section_id: str, style: int = 1) -> str:
        return self.call("create_new_page", section_id=section_id, new_page_style=style)["page_id"]
