"""One object that owns every call to OneNote: bridge, hierarchy cache, identifier resolution, page I/O."""

from __future__ import annotations

import html
import re
import threading
import time
import xml.etree.ElementTree as ET
from typing import Any, Callable

from .bridge import OneNoteBridge, OneNoteBridgeError
from .constants import HIERARCHY_SCOPES, PAGE_INFO, XML_SCHEMA_2013, enum_value
from .hierarchy import ONENOTE_ID, Item, in_recycle_bin, normal_path, of_type, parse_hierarchy, parse_xml, resolve_item, without_recycle_bin

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
# OneNote's own HRESULTs (Microsoft "Error codes (OneNote)"); the COM message for them is only the number.
ONENOTE_ERRORS = {
    0x80042000: "OneNote rejected the XML as malformed.",
    0x80042001: "OneNote rejected the XML as invalid.",
    0x80042002: "The section could not be created.",
    0x80042003: "The section could not be opened.",
    0x80042004: "The section does not exist.",
    0x80042005: "The page does not exist.",
    0x80042006: "The file does not exist.",
    0x80042007: "The image could not be inserted.",
    0x80042008: "The ink could not be inserted.",
    0x80042009: "The HTML could not be inserted.",
    0x8004200B: "The section is read-only.",
    0x8004200C: "The page is read-only (pages in the recycle bin cannot be changed).",
    0x8004200D: "The outline text could not be inserted.",
    0x8004200E: "The page object does not exist (it may have been deleted or the ID is stale).",
    0x8004200F: "The binary object does not exist.",
    0x80042010: "The page changed since it was read; read it again and retry.",
    0x80042011: "The section group does not exist.",
    0x80042012: "The page does not exist in that section group.",
    0x80042014: "The object does not exist.",
    0x80042015: "The notebook does not exist.",
    0x80042016: "The file could not be inserted.",
    0x80042017: "The name is not allowed (too long, or it holds characters OneNote does not accept).",
    0x80042018: "The folder does not exist.",
    0x80042019: "The search query is invalid.",
    0x8004201A: "A file with that name already exists.",
    0x8004201B: "The section is password-protected and locked; unlock it in OneNote first.",
    0x8004201C: "This action is disabled by policy.",
    0x8004201D: "OneNote has not finished syncing this content yet; try again shortly.",
    0x8004201E: "This is an old-format (2007) section; convert it in OneNote first.",
    0x8004201F: "The merge failed.",
    0x80042020: "The XML does not match OneNote's schema.",
    0x80042022: "OneNote refused the change because content would be lost (newer content features on the page).",
    0x80042023: "OneNote timed out.",
    0x80042024: "A recording is in progress.",
    0x8004202C: "The selection is invalid.",
    0x8004202D: "The conversion failed.",
    0x8004202E: "Items in the recycle bin cannot be changed; restore them first.",
    0x80042030: "OneNote has a dialog open; close it and retry.",
}


def _page_id_of(xml: str) -> str:
    match = re.search(r"<(?:\w+:)?Page\b[^>]*?\sID=\"([^\"]+)\"", xml)
    return html.unescape(match.group(1)) if match else ""


class OneNote:
    def __init__(self, bridge: OneNoteBridge, cache_seconds: float = 5.0) -> None:
        self.bridge = bridge
        self.cache_seconds = cache_seconds
        self._lock = threading.Lock()
        self._cache: dict[tuple[str, str], tuple[float, list[Item]]] = {}
        self.on_hierarchy_read: Callable[[], None] | None = None
        self.on_page_changed: Callable[[str], None] | None = None

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
            explained = ONENOTE_ERRORS.get((exc.hresult or 0) & 0xFFFFFFFF)
            if explained:
                raise RuntimeError(f"{explained} (OneNote error 0x{exc.hresult & 0xFFFFFFFF:08X} in {operation})") from exc
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
        self.ensure_window()
        self.call("update_hierarchy", xml=ET.tostring(tree, encoding="unicode"), schema=XML_SCHEMA_2013)

    def ensure_window(self) -> None:
        """Make sure OneNote has a window: a hierarchy update crashes OneNote when only COM started it."""

        if not self.call("ensure_window", timeout_seconds=20).get("window"):
            raise RuntimeError(
                "OneNote has no open window and could not be started with one; open OneNote and retry. "
                "(Changing the page or section order without a window crashes OneNote.)"
            )

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
        """One object by ID, with its full path and parent, without reading every page of every notebook."""

        try:
            items = parse_hierarchy(self.hierarchy_xml(object_id, "self"))
        except RuntimeError:
            return None
        return self._located(items[0]) if items else None

    def _located(self, item: Item) -> Item:
        """Fill path, level and parent of an item read on its own (that read only knows the item itself)."""

        if item["type"] == "notebook":
            return item
        if item["type"] == "page":
            # A page ID starts with the GUID of the section that holds it.
            guid = item["id"].split("}{", 1)[0]
            for fresh in (False, True):
                section = next(
                    (s for s in self.items("", "sections", fresh) if s["type"] == "section" and s["id"].split("}{", 1)[0] == guid),
                    None,
                )
                if section:
                    return {
                        **item,
                        "path": f"{section['path']}/{item['name']}",
                        "level": section["level"] + 1,
                        "parent_id": section["id"],
                        "notebook_name": section["notebook_name"],
                        "section_name": section["name"],
                    }
        found = self.find(id=item["id"], item_type=item["type"]) or self.find(id=item["id"], item_type=item["type"], fresh=True)
        return {**item, **found} if found else item

    def resolve(self, identifier: str, item_type: str | None = None) -> Item:
        """ID, exact path, or unique name -> live item, with its path and parent."""

        identifier = (identifier or "").strip()
        if not identifier:
            raise ValueError("An identifier is required.")
        if ONENOTE_ID.match(identifier):
            item = self.lookup(identifier)
            if item:
                if item_type and item["type"] != item_type:
                    raise ValueError(f"'{identifier}' is a {item['type']}, expected a {item_type}.")
                return item
        scope = "sections" if item_type in CONTAINER_TYPES else "pages"
        return resolve_item(self.items("", scope), identifier, item_type)

    def live(self, identifier: str, item_type: str | None = None) -> Item:
        """resolve(), refusing items in the recycle bin (OneNote does not let them be changed)."""

        item = self.resolve(identifier, item_type)
        if in_recycle_bin(item):
            raise ValueError(f"'{item['path']}' is in the recycle bin; restore it in OneNote first.")
        return item

    def find(self, *, id: str = "", path: str = "", item_type: str | None = None, fresh: bool = False) -> Item | None:
        scope = "sections" if item_type in CONTAINER_TYPES else "pages"
        for item in self.items("", scope, fresh=fresh):
            if item_type and item["type"] != item_type:
                continue
            if (id and item["id"] == id) or (path and normal_path(item["path"]) == normal_path(path)):
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
        root = xml if isinstance(xml, ET.Element) else None
        if root is not None:
            xml = ET.tostring(root, encoding="unicode")
        try:
            self.call("update_page_content", xml=xml, schema=XML_SCHEMA_2013, force=force)
        finally:
            page_id = root.attrib.get("ID", "") if root is not None else _page_id_of(xml)
            if page_id and self.on_page_changed:
                self.on_page_changed(page_id)

    def delete_object(self, page_id: str, object_id: str) -> None:
        try:
            self.call("delete_page_content", page_id=page_id, object_id=object_id, force=True)
        finally:
            if self.on_page_changed:
                self.on_page_changed(page_id)

    def new_page(self, section_id: str, style: int = 1) -> str:
        return self.call("create_new_page", section_id=section_id, new_page_style=style)["page_id"]
