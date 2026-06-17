"""MCP server exposing a pure-local Microsoft OneNote control surface."""

from __future__ import annotations

import base64
import os
import re
import sys
import time
from pathlib import Path
from typing import Any

from mcp.server.fastmcp import FastMCP

from .bridge import OneNoteBridge, OneNoteBridgeError
from .constants import (
    CREATE_FILE_TYPES,
    FILING_LOCATION_TYPES,
    FILING_LOCATIONS,
    HIERARCHY_SCOPES,
    NEW_PAGE_STYLES,
    PAGE_INFO,
    PUBLISH_FORMATS,
    SPECIAL_LOCATIONS,
    XML_SCHEMA_2013,
)
from .image_utils import proportional_dimensions
from .xml_utils import (
    build_image_page_update_xml,
    build_page_update_xml,
    collect_page_objects,
    DELETABLE_PAGE_OBJECT_TYPES,
    filter_items,
    parse_hierarchy,
    resolve_item,
    text_from_page_xml,
    title_from_page_xml,
)


MCP_NAME = "local-onenote"
DEFAULT_TIMEOUT = int(os.environ.get("LOCAL_ONENOTE_MCP_TIMEOUT", "90"))
MAX_TEXT_CHARS = int(os.environ.get("LOCAL_ONENOTE_MCP_MAX_TEXT_CHARS", "60000"))

mcp = FastMCP(MCP_NAME)
bridge = OneNoteBridge(timeout_seconds=DEFAULT_TIMEOUT)


def _error(message: str) -> dict[str, Any]:
    return {"ok": False, "error": message}


def _ok(**data: Any) -> dict[str, Any]:
    return {"ok": True, **data}


def _enum(name: str, value: str, options: dict[str, int]) -> int:
    key = value.casefold()
    if key not in options:
        allowed = ", ".join(sorted(options))
        raise ValueError(f"{name} must be one of: {allowed}")
    return options[key]


def _bridge(operation: str, **params: Any) -> dict[str, Any]:
    try:
        return bridge.call(operation, **params)
    except OneNoteBridgeError as exc:
        raise RuntimeError(str(exc)) from exc


def _hierarchy_xml(start_id: str = "", scope: str = "pages") -> str:
    return _bridge(
        "get_hierarchy",
        start_id=start_id,
        scope=_enum("scope", scope, HIERARCHY_SCOPES),
        schema=XML_SCHEMA_2013,
    )["xml"]


def _hierarchy_items(start_id: str = "", scope: str = "pages") -> list[dict[str, Any]]:
    return parse_hierarchy(_hierarchy_xml(start_id, scope))


def _resolve_id(identifier: str, item_type: str | None = None) -> str:
    if not identifier:
        return ""
    items = _hierarchy_items("", "pages")
    return resolve_item(items, identifier, item_type)["id"]


def _resolve_item(identifier: str, item_type: str | None = None) -> dict[str, Any]:
    items = _hierarchy_items("", "pages")
    return resolve_item(items, identifier, item_type)


def _find_item_by_path(path: str, item_type: str | None = None) -> dict[str, Any] | None:
    target = path.casefold()
    for item in _hierarchy_items("", "pages"):
        if item_type and item.get("type") != item_type:
            continue
        if item.get("path", "").casefold() == target:
            return item
    return None


def _find_item_by_id(object_id: str, item_type: str | None = None) -> dict[str, Any] | None:
    if not object_id:
        return None
    for item in _hierarchy_items("", "pages"):
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
    for attempt in range(retries):
        item = _find_item_by_path(expected_path, item_type)
        if item:
            return item
        item = _find_item_by_id(fallback_id, item_type)
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


def _page_xml(page_id: str, page_info: str = "basic") -> str:
    return _bridge(
        "get_page_content",
        page_id=page_id,
        page_info=_enum("page_info", page_info, PAGE_INFO),
        schema=XML_SCHEMA_2013,
    )["xml"]


def _local_text_search(
    start_id: str,
    query: str,
    max_results: int,
    include_recycle_bin: bool,
) -> list[dict[str, Any]]:
    items = _hierarchy_items(start_id, "pages")
    pages = filter_items(items, "page")
    if not include_recycle_bin:
        pages = _without_recycle_bin(pages)
    query_lower = query.casefold()
    matches = []
    for page in pages:
        if len(matches) >= max(1, max_results):
            break
        haystacks = [page.get("name", ""), page.get("path", "")]
        try:
            haystacks.append(text_from_page_xml(_page_xml(page["id"], "basic")))
        except Exception as exc:
            page["scan_error"] = str(exc)
        if any(query_lower in value.casefold() for value in haystacks if value):
            matches.append(page)
    return matches


REPLACE_BODY_OBJECT_TYPES = {"Outline", "Image", "InkDrawing", "FileAttachment", "InsertedFile", "MediaFile"}
IDENTIFIER_RESOLUTION_ORDER = ["id", "exact_path", "unique_name"]
IDENTIFIER_TYPES = {"notebook", "section_group", "section", "page"}


@mcp.tool()
async def health_check() -> dict[str, Any]:
    """Verify local OneNote COM access and return a small hierarchy summary."""

    try:
        items = _without_recycle_bin(_hierarchy_items("", "sections"))
        notebooks = filter_items(items, "notebook")
        sections = filter_items(items, "section")
        return _ok(
            server=MCP_NAME,
            transport="stdio",
            python_executable=sys.executable,
            module_path=str(Path(__file__).resolve()),
            process_cwd=str(Path.cwd()),
            timeout_seconds=DEFAULT_TIMEOUT,
            max_text_chars=MAX_TEXT_CHARS,
            identifier_resolution_order=IDENTIFIER_RESOLUTION_ORDER,
            search_default_backend="local_scan",
            content_formats=["plain", "html", "markdown"],
            notebooks=len(notebooks),
            sections=len(sections),
            write_backend="OneNote desktop COM API",
        )
    except Exception as exc:
        return _error(str(exc))


@mcp.tool()
async def resolve_identifier(identifier: str, item_type: str = "") -> dict[str, Any]:
    """Resolve a OneNote identifier to one live object before using it in another tool."""

    try:
        if not identifier:
            raise ValueError("identifier is required.")
        normalized_type = item_type.strip().casefold() or None
        if normalized_type and normalized_type not in IDENTIFIER_TYPES:
            allowed = ", ".join(sorted(IDENTIFIER_TYPES))
            raise ValueError(f"item_type must be empty or one of: {allowed}")
        item = _resolve_item(identifier, normalized_type)
        return _ok(item=item, identifier_resolution_order=IDENTIFIER_RESOLUTION_ORDER)
    except Exception as exc:
        return _error(str(exc))


@mcp.tool()
async def get_special_locations() -> dict[str, Any]:
    """Return OneNote's local special folders: backup, unfiled, and default notebook folder."""

    try:
        locations = {}
        for name, value in SPECIAL_LOCATIONS.items():
            locations[name] = _bridge("get_special_location", location=value)["path"]
        return _ok(locations=locations)
    except Exception as exc:
        return _error(str(exc))


@mcp.tool()
async def list_hierarchy(
    start_identifier: str = "",
    scope: str = "pages",
    include_xml: bool = False,
    include_recycle_bin: bool = False,
) -> dict[str, Any]:
    """List live OneNote hierarchy objects. Identifiers may be an ID, exact path, or unique name."""

    try:
        start_id = _resolve_id(start_identifier) if start_identifier else ""
        xml = _hierarchy_xml(start_id, scope)
        items = parse_hierarchy(xml)
        if not include_recycle_bin:
            items = _without_recycle_bin(items)
        data = _ok(items=items, count=len(items))
        if include_xml:
            data["xml"] = xml
        return data
    except Exception as exc:
        return _error(str(exc))


@mcp.tool()
async def list_notebooks() -> dict[str, Any]:
    """List live notebooks currently known to the local OneNote desktop app."""

    try:
        items = _hierarchy_items("", "notebooks")
        notebooks = filter_items(items, "notebook")
        return _ok(notebooks=notebooks, count=len(notebooks))
    except Exception as exc:
        return _error(str(exc))


@mcp.tool()
async def list_sections(notebook_identifier: str = "", include_recycle_bin: bool = False) -> dict[str, Any]:
    """List sections, optionally restricted to a notebook ID, exact path, or unique name."""

    try:
        items = _hierarchy_items("", "sections")
        if not include_recycle_bin:
            items = _without_recycle_bin(items)
        sections = filter_items(items, "section")
        if notebook_identifier:
            notebook = resolve_item(items, notebook_identifier, "notebook")
            prefix = notebook["path"] + "/"
            sections = [section for section in sections if section["path"].startswith(prefix)]
        return _ok(sections=sections, count=len(sections))
    except Exception as exc:
        return _error(str(exc))


@mcp.tool()
async def list_pages(section_identifier: str, include_xml: bool = False, include_recycle_bin: bool = False) -> dict[str, Any]:
    """List pages in a section selected by section ID, exact path, or unique name."""

    try:
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
    except Exception as exc:
        return _error(str(exc))


@mcp.tool()
async def get_page(page_identifier: str, include_xml: bool = False, page_info: str = "basic", max_chars: int = MAX_TEXT_CHARS) -> dict[str, Any]:
    """Read a page's title, plain text, object IDs, and optionally raw OneNote XML."""

    try:
        page = _resolve_item(page_identifier, "page")
        xml = _page_xml(page["id"], page_info)
        text = _truncate(text_from_page_xml(xml), max_chars)
        data = _ok(
            page=page,
            title=title_from_page_xml(xml),
            text=text,
            objects=collect_page_objects(xml),
        )
        if include_xml:
            data["xml"] = xml
        return data
    except Exception as exc:
        return _error(str(exc))


@mcp.tool()
async def get_page_xml(page_identifier: str, page_info: str = "basic") -> dict[str, Any]:
    """Return raw OneNote XML for a page."""

    try:
        page_id = _resolve_id(page_identifier, "page")
        return _ok(xml=_page_xml(page_id, page_info))
    except Exception as exc:
        return _error(str(exc))


@mcp.tool()
async def get_page_text(page_identifier: str, max_chars: int = MAX_TEXT_CHARS) -> dict[str, Any]:
    """Return plain text extracted from a OneNote page."""

    try:
        page_id = _resolve_id(page_identifier, "page")
        text = text_from_page_xml(_page_xml(page_id, "basic"))
        return _ok(text=_truncate(text, max_chars), chars=len(text))
    except Exception as exc:
        return _error(str(exc))


@mcp.tool()
async def get_page_objects(page_identifier: str) -> dict[str, Any]:
    """List page content objects such as outlines, images, attachments, and callback IDs."""

    try:
        page_id = _resolve_id(page_identifier, "page")
        objects = collect_page_objects(_page_xml(page_id, "all"))
        return _ok(objects=objects, count=len(objects))
    except Exception as exc:
        return _error(str(exc))


@mcp.tool()
async def get_binary_content(page_identifier: str, callback_id: str) -> dict[str, Any]:
    """Read binary page content by callback ID returned from get_page_objects."""

    try:
        page_id = _resolve_id(page_identifier, "page")
        result = _bridge("get_binary_page_content", page_id=page_id, callback_id=callback_id)
        return _ok(base64=result["base64"])
    except Exception as exc:
        return _error(str(exc))


@mcp.tool()
async def search_pages(
    query: str,
    start_identifier: str = "",
    max_results: int = 20,
    include_snippets: bool = True,
    include_unindexed: bool = True,
    include_recycle_bin: bool = False,
) -> dict[str, Any]:
    """Search pages. include_unindexed=true scans live page text instead of relying on OneNote's index."""

    try:
        start_id = _resolve_id(start_identifier) if start_identifier else ""
        used_local_scan = include_unindexed
        if include_unindexed:
            pages = _local_text_search(start_id, query, max_results, include_recycle_bin)
        else:
            try:
                xml = _bridge(
                    "find_pages",
                    start_id=start_id,
                    query=query,
                    include_unindexed=False,
                    display=False,
                    schema=XML_SCHEMA_2013,
                )["xml"]
                pages = filter_items(parse_hierarchy(xml), "page")
            except Exception:
                used_local_scan = True
                pages = _local_text_search(start_id, query, max_results, include_recycle_bin)
        if not include_recycle_bin:
            pages = _without_recycle_bin(pages)
        pages = pages[: max(1, max_results)]
        if include_snippets:
            q = query.casefold()
            for page in pages:
                try:
                    text = text_from_page_xml(_page_xml(page["id"], "basic"))
                    idx = text.casefold().find(q)
                    if idx >= 0:
                        start = max(0, idx - 160)
                        end = min(len(text), idx + len(query) + 240)
                        page["snippet"] = text[start:end].strip()
                except Exception as exc:
                    page["snippet_error"] = str(exc)
        return _ok(pages=pages, count=len(pages), search_backend="local_scan" if used_local_scan else "onenote_index")
    except Exception as exc:
        return _error(str(exc))


@mcp.tool()
async def find_meta(start_identifier: str, name: str, include_unindexed: bool = True) -> dict[str, Any]:
    """Find pages or objects with matching OneNote meta name."""

    try:
        start_id = _resolve_id(start_identifier) if start_identifier else ""
        xml = _bridge(
            "find_meta",
            start_id=start_id,
            name=name,
            include_unindexed=include_unindexed,
            schema=XML_SCHEMA_2013,
        )["xml"]
        items = parse_hierarchy(xml)
        return _ok(items=items, count=len(items), xml=xml)
    except Exception as exc:
        return _error(str(exc))


@mcp.tool()
async def get_hyperlink(object_identifier: str, page_content_object_id: str = "", web: bool = False) -> dict[str, Any]:
    """Return a OneNote client or web hyperlink for an object."""

    try:
        object_id = _resolve_id(object_identifier)
        operation = "get_web_hyperlink" if web else "get_hyperlink"
        result = _bridge(operation, object_id=object_id, page_content_object_id=page_content_object_id)
        return _ok(hyperlink=result["hyperlink"])
    except Exception as exc:
        return _error(str(exc))


@mcp.tool()
async def get_parent(object_identifier: str) -> dict[str, Any]:
    """Return the parent object ID for a notebook hierarchy object."""

    try:
        object_id = _resolve_id(object_identifier)
        return _ok(parent_id=_bridge("get_hierarchy_parent", object_id=object_id)["parent_id"])
    except Exception as exc:
        return _error(str(exc))


@mcp.tool()
async def open_hierarchy(path: str, relative_to_identifier: str = "", create_type: str = "none") -> dict[str, Any]:
    """Open or create a notebook, section group, or section. Existing OneNote hierarchy paths resolve directly."""

    try:
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
        return data
    except Exception as exc:
        return _error(str(exc))


@mcp.tool()
async def create_notebook(name_or_path: str, base_folder: str = "") -> dict[str, Any]:
    """Create a local notebook folder and open it in OneNote."""

    try:
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
        return _ok(path=str(notebook_path), notebook_id=result["object_id"])
    except Exception as exc:
        return _error(str(exc))


@mcp.tool()
async def create_section(parent_identifier: str, section_name: str) -> dict[str, Any]:
    """Create a section under a notebook or section group."""

    try:
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
    except Exception as exc:
        return _error(str(exc))


@mcp.tool()
async def create_section_group(parent_identifier: str, group_name: str) -> dict[str, Any]:
    """Create a section group under a notebook or another section group."""

    try:
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
    except Exception as exc:
        return _error(str(exc))


@mcp.tool()
async def create_page(
    section_identifier: str,
    title: str,
    content: str = "",
    content_format: str = "plain",
    new_page_style: str = "blank_with_title",
) -> dict[str, Any]:
    """Create a page in a local OneNote section. content_format accepts plain, html, or markdown."""

    try:
        section = _resolve_item(section_identifier, "section")
        result = _bridge(
            "create_new_page",
            section_id=section["id"],
            new_page_style=_enum("new_page_style", new_page_style, NEW_PAGE_STYLES),
        )
        page_id = result["page_id"]
        xml = build_page_update_xml(page_id, title=title, content=content, content_format=content_format)
        _bridge("update_page_content", xml=xml, schema=XML_SCHEMA_2013, force=False)
        expected_path = _friendly_child_path(section["path"], title)
        page = _refresh_created_item(expected_path=expected_path, item_type="page", fallback_id=page_id)
        return _ok(page_id=page["id"] if page else page_id, page=page, section=section, title=title, path=expected_path)
    except Exception as exc:
        return _error(str(exc))


@mcp.tool()
async def update_page_title(page_identifier: str, title: str) -> dict[str, Any]:
    """Update a page title."""

    try:
        page_id = _resolve_id(page_identifier, "page")
        xml = build_page_update_xml(page_id, title=title)
        _bridge("update_page_content", xml=xml, schema=XML_SCHEMA_2013, force=False)
        return _ok(page_id=page_id, title=title)
    except Exception as exc:
        return _error(str(exc))


@mcp.tool()
async def append_to_page(
    page_identifier: str,
    content: str,
    content_format: str = "plain",
    x: float | None = None,
    y: float | None = None,
) -> dict[str, Any]:
    """Append a new outline block to a page. content_format accepts plain, html, or markdown."""

    try:
        page_id = _resolve_id(page_identifier, "page")
        xml = build_page_update_xml(page_id, content=content, content_format=content_format, x=x, y=y)
        _bridge("update_page_content", xml=xml, schema=XML_SCHEMA_2013, force=False)
        return _ok(page_id=page_id, appended=True)
    except Exception as exc:
        return _error(str(exc))


@mcp.tool()
async def add_image_to_page(
    page_identifier: str,
    image_path: str,
    image_format: str = "",
    x: float = 36.0,
    y: float = 120.0,
    width: float | None = None,
    height: float | None = None,
) -> dict[str, Any]:
    """Add a local image file to a OneNote page."""

    try:
        path = Path(image_path)
        if not path.is_file():
            raise ValueError(f"Image file not found: {image_path}")
        fmt = image_format or path.suffix.lstrip(".")
        if not fmt:
            raise ValueError("image_format is required when image_path has no extension.")
        resolved_width, resolved_height = proportional_dimensions(path, width, height)
        image_base64 = base64.b64encode(path.read_bytes()).decode("ascii")
        page_id = _resolve_id(page_identifier, "page")
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
    except Exception as exc:
        return _error(str(exc))


@mcp.tool()
async def replace_page_body(
    page_identifier: str,
    content: str,
    title: str | None = None,
    content_format: str = "plain",
) -> dict[str, Any]:
    """Delete existing page content objects and write new body content. content_format accepts plain, html, or markdown."""

    try:
        page_id = _resolve_id(page_identifier, "page")
        page_xml = _page_xml(page_id, "all")
        objects = collect_page_objects(page_xml)
        deleted = []
        for obj in objects:
            if obj.get("type") not in REPLACE_BODY_OBJECT_TYPES:
                continue
            object_id = obj.get("object_id")
            if not object_id:
                continue
            _bridge("delete_page_content", page_id=page_id, object_id=object_id, force=True)
            deleted.append(object_id)
        xml = build_page_update_xml(page_id, title=title, content=content, content_format=content_format)
        _bridge("update_page_content", xml=xml, schema=XML_SCHEMA_2013, force=True)
        return _ok(page_id=page_id, deleted_objects=deleted, replaced=True)
    except Exception as exc:
        return _error(str(exc))


@mcp.tool()
async def delete_page_content(page_identifier: str, object_id: str) -> dict[str, Any]:
    """Delete one deletable page content object by object ID. Use get_page_objects to find delete_supported objects."""

    try:
        page_id = _resolve_id(page_identifier, "page")
        objects = collect_page_objects(_page_xml(page_id, "all"))
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
    except Exception as exc:
        return _error(str(exc))


@mcp.tool()
async def delete_hierarchy(object_identifier: str, permanently: bool = False) -> dict[str, Any]:
    """Delete a notebook, section group, section, or page."""

    try:
        item = _resolve_item(object_identifier)
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
    except Exception as exc:
        return _error(str(exc))


@mcp.tool()
async def update_page_xml(xml: str, force: bool = False) -> dict[str, Any]:
    """Advanced: submit raw OneNote page XML to UpdatePageContent."""

    try:
        _bridge("update_page_content", xml=xml, schema=XML_SCHEMA_2013, force=force)
        return _ok(updated=True, force=force)
    except Exception as exc:
        return _error(str(exc))


@mcp.tool()
async def update_hierarchy_xml(xml: str) -> dict[str, Any]:
    """Advanced: submit raw OneNote hierarchy XML to UpdateHierarchy."""

    try:
        _bridge("update_hierarchy", xml=xml, schema=XML_SCHEMA_2013)
        return _ok(updated=True)
    except Exception as exc:
        return _error(str(exc))


@mcp.tool()
async def publish_object(object_identifier: str, target_path: str, format: str = "pdf", overwrite: bool = False) -> dict[str, Any]:
    """Export a notebook, section, or page to a local file."""

    try:
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
    except Exception as exc:
        return _error(str(exc))


@mcp.tool()
async def navigate_to(object_identifier: str, page_content_object_id: str = "", new_window: bool = False) -> dict[str, Any]:
    """Open a OneNote object in the desktop app."""

    try:
        object_id = _resolve_id(object_identifier)
        _bridge("navigate_to", object_id=object_id, page_content_object_id=page_content_object_id, new_window=new_window)
        return _ok(navigated=True)
    except Exception as exc:
        return _error(str(exc))


@mcp.tool()
async def navigate_to_url(url: str, new_window: bool = False) -> dict[str, Any]:
    """Open a OneNote URL in the desktop app."""

    try:
        _bridge("navigate_to_url", url=url, new_window=new_window)
        return _ok(navigated=True)
    except Exception as exc:
        return _error(str(exc))


@mcp.tool()
async def sync_hierarchy(object_identifier: str) -> dict[str, Any]:
    """Ask OneNote to sync a notebook hierarchy object."""

    try:
        object_id = _resolve_id(object_identifier)
        _bridge("sync_hierarchy", object_id=object_id)
        return _ok(object_id=object_id, synced=True)
    except Exception as exc:
        return _error(str(exc))


@mcp.tool()
async def close_notebook(notebook_identifier: str, force: bool = False) -> dict[str, Any]:
    """Close a notebook in the desktop OneNote app."""

    try:
        notebook_id = _resolve_id(notebook_identifier, "notebook")
        _bridge("close_notebook", notebook_id=notebook_id, force=force)
        return _ok(notebook_id=notebook_id, closed=True)
    except Exception as exc:
        return _error(str(exc))


@mcp.tool()
async def merge_sections(source_section_identifier: str, destination_section_identifier: str) -> dict[str, Any]:
    """Merge one section into another."""

    try:
        source_id = _resolve_id(source_section_identifier, "section")
        destination_id = _resolve_id(destination_section_identifier, "section")
        _bridge("merge_sections", source_section_id=source_id, destination_section_id=destination_id)
        return _ok(source_section_id=source_id, destination_section_id=destination_id, merged=True)
    except Exception as exc:
        return _error(str(exc))


@mcp.tool()
async def set_filing_location(filing_location: str, filing_location_type: str, section_or_page_identifier: str) -> dict[str, Any]:
    """Set OneNote's local filing location for email, web clips, printouts, and similar content."""

    try:
        object_id = _resolve_id(section_or_page_identifier)
        _bridge(
            "set_filing_location",
            filing_location=_enum("filing_location", filing_location, FILING_LOCATIONS),
            filing_location_type=_enum("filing_location_type", filing_location_type, FILING_LOCATION_TYPES),
            section_or_page_id=object_id,
        )
        return _ok(object_id=object_id, updated=True)
    except Exception as exc:
        return _error(str(exc))


def main() -> None:
    mcp.run(transport="stdio")


if __name__ == "__main__":
    main()
