"""Read-only tools: discovery, page reading, images, search."""

from __future__ import annotations

import base64
import sys
import tempfile
import uuid
from pathlib import Path
from typing import Any

from mcp.server.fastmcp.utilities.types import Image

from .. import __version__
from ..app import READ_ONLY, backups, indexer, ok, onenote, settings, text_index, tool
from ..backup import FILE_TAGS
from ..constants import PUBLISH_FORMATS, SPECIAL_LOCATIONS, XML_SCHEMA_2013
from ..editing import PageEditor
from ..hierarchy import of_type, parse_hierarchy, parse_xml, resolve_item, without_recycle_bin
from ..pages import file_bytes
from ..pagexml import collect_page_objects
from ..reader import page_title, render_page
from .. import search

MAX_IMAGE_BYTES = 3_500_000
IDENTIFIER_TYPES = {"notebook", "section_group", "section", "page"}
SEARCH_BACKENDS = {"auto", "index", "local_scan"}


def _truncate(text: str, max_chars: int) -> str:
    if max_chars <= 0 or len(text) <= max_chars:
        return text
    return text[:max_chars] + f"\n\n[truncated: {len(text) - max_chars} chars omitted]"


def _read_format(value: str) -> str:
    key = (value or "text").strip().casefold()
    if key in {"plain", "text"}:
        return "text"
    if key in {"markdown", "md"}:
        return "markdown"
    raise ValueError("format must be 'text' or 'markdown'.")


@tool(READ_ONLY)
def health_check() -> dict[str, Any]:
    """Check OneNote access; report version, features, special folders, text-index and backup status."""

    items = without_recycle_bin(onenote.items("", "sections"))
    locations = {name: onenote.call("get_special_location", location=value)["path"] for name, value in SPECIAL_LOCATIONS.items()}
    return ok(
        server="local-onenote",
        version=__version__,
        python_executable=sys.executable,
        module_path=str(Path(__file__).resolve().parent.parent),
        notebooks=len(of_type(items, "notebook")),
        sections=len(of_type(items, "section")),
        special_locations=locations,
        content_formats=["plain", "html", "markdown"],
        read_formats=["text", "markdown"],
        timeout_seconds=settings.timeout_seconds,
        text_index=indexer.status(),
        backups=backups.stats(),
    )


@tool(READ_ONLY)
def resolve_identifier(identifier: str, item_type: str = "") -> dict[str, Any]:
    """Resolve an ID, exact path ("Notebook/Section/Page") or unique name to one live object."""

    kind = item_type.strip().casefold() or None
    if kind and kind not in IDENTIFIER_TYPES:
        raise ValueError(f"item_type must be empty or one of: {', '.join(sorted(IDENTIFIER_TYPES))}")
    return ok(item=onenote.resolve(identifier, kind, full=True))


@tool(READ_ONLY)
def list_notebooks() -> dict[str, Any]:
    """List the notebooks open in OneNote."""

    notebooks = of_type(onenote.items("", "notebooks"), "notebook")
    return ok(notebooks=notebooks, count=len(notebooks))


@tool(READ_ONLY)
def list_sections(notebook_identifier: str = "", include_recycle_bin: bool = False) -> dict[str, Any]:
    """List sections (with their section-group paths), optionally only those in one notebook."""

    items = onenote.items("", "sections")
    if not include_recycle_bin:
        items = without_recycle_bin(items)
    sections = of_type(items, "section")
    if notebook_identifier:
        prefix = resolve_item(items, notebook_identifier, "notebook")["path"] + "/"
        sections = [section for section in sections if section["path"].startswith(prefix)]
    return ok(sections=sections, count=len(sections))


@tool(READ_ONLY)
def list_pages(section_identifier: str, include_recycle_bin: bool = False) -> dict[str, Any]:
    """List the pages of a section in order. pageLevel 2/3 marks subpages."""

    section = onenote.resolve(section_identifier, "section", full=True)
    pages = of_type(parse_hierarchy(onenote.hierarchy_xml(section["id"], "pages")), "page")
    if not include_recycle_bin:
        pages = without_recycle_bin(pages)
    for page in pages:
        page["path"] = f"{section['path']}/{page['name']}"
        page["parent_id"] = section["id"]
        page["notebook_name"] = section.get("notebook_name")
        page["section_name"] = section.get("name")
    return ok(section=section, pages=pages, count=len(pages))


@tool(READ_ONLY)
def list_hierarchy(
    start_identifier: str = "",
    scope: str = "pages",
    include_recycle_bin: bool = False,
    max_items: int = 500,
) -> dict[str, Any]:
    """List notebooks, section groups, sections and (scope="pages") pages below a starting point.

    max_items caps the list (0 = no cap); use scope="sections" or a start_identifier for large collections.
    """

    start_id = onenote.resolve(start_identifier)["id"] if start_identifier else ""
    items = onenote.items(start_id, scope)
    if not include_recycle_bin:
        items = without_recycle_bin(items)
    total = len(items)
    if max_items > 0:
        items = items[:max_items]
    return ok(items=items, count=len(items), total=total, truncated=len(items) < total)


@tool(READ_ONLY)
def get_page(
    page_identifier: str,
    format: str = "text",
    include_ocr: bool = True,
    include_paragraphs: bool = False,
    include_objects: bool = False,
    include_xml: bool = False,
    max_chars: int = settings.max_text_chars,
) -> dict[str, Any]:
    """Read a page's title and body.

    format="text" gives plain text; format="markdown" keeps headings, bold/italic, links, lists, to-dos and tables.
    include_ocr adds the text OneNote recognised inside images.
    include_paragraphs lists every paragraph with its id, text, depth, list type, style and to-do state; use those ids
    with edit_page. include_objects lists outlines, images, ink and attachments with object_id (for
    delete_page_content) and callback_id (for get_page_image / get_binary_content). include_xml adds the raw XML.
    """

    fmt = _read_format(format)
    page = onenote.resolve(page_identifier, "page")
    xml = onenote.page_xml(page["id"])
    root = parse_xml(xml)
    text = render_page(xml, fmt, include_ocr=include_ocr, include_title=False)
    data = ok(page=page, title=page_title(xml), format=fmt, text=_truncate(text, max_chars), chars=len(text))
    if include_paragraphs:
        data["paragraphs"] = PageEditor(root).paragraphs()
    if include_objects:
        data["objects"] = collect_page_objects(root)
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


def _binary(page_identifier: str, callback_id: str = "", file_name: str = "") -> bytes:
    page_id = onenote.resolve(page_identifier, "page")["id"]
    if callback_id:
        return base64.b64decode(onenote.call("get_binary_page_content", page_id=page_id, callback_id=callback_id)["base64"])
    if not file_name:
        raise ValueError("Give callback_id (images, ink) or file_name (attachments).")
    for node in parse_xml(onenote.page_xml(page_id)).iter():
        if node.tag in FILE_TAGS and node.attrib.get("preferredName") == file_name:
            data = file_bytes(page_id, node)
            if data is not None:
                return data
    raise ValueError(f"No readable attachment named '{file_name}' on this page.")


@tool(READ_ONLY, structured=False)
def get_page_image(page_identifier: str, callback_id: str) -> Any:
    """Show an image from a page to the model. Take callback_id from get_page(include_objects=true)."""

    raw = _binary(page_identifier, callback_id)
    fmt = _image_format(raw)
    if fmt is None:
        raise ValueError("This object is not a PNG, JPEG, GIF or WebP image; use get_binary_content to download it.")
    if len(raw) > MAX_IMAGE_BYTES:
        raise ValueError(f"Image is {len(raw) // 1024} KB, above the {MAX_IMAGE_BYTES // 1024} KB inline limit; use get_binary_content.")
    return Image(data=raw, format=fmt)


@tool(READ_ONLY, structured=False)
def view_page(page_identifier: str, max_pages: int = 3, width: int = 1000) -> Any:
    """Show a page to the model as it looks in OneNote, including handwriting, drawings and layout.

    The page is printed to PDF and each printed (A4) page becomes one JPEG image, width pixels wide.
    Use it for pages whose ink or layout get_page cannot put into text.
    """

    page_id = onenote.resolve(page_identifier, "page")["id"]
    pdf = Path(tempfile.gettempdir()) / f"local-onenote-mcp-{uuid.uuid4().hex}.pdf"
    try:
        onenote.call("publish", object_id=page_id, target_path=str(pdf), format=PUBLISH_FORMATS["pdf"])
        rendered = onenote.call("render_pdf", path=str(pdf), width=min(max(width, 200), 2000), max_pages=min(max(max_pages, 1), 10))
    finally:
        pdf.unlink(missing_ok=True)
    images = rendered.get("images") or []
    if isinstance(images, str):
        images = [images]
    shown = f"pages 1-{len(images)}" if len(images) > 1 else "page 1"
    return [f"Printed {shown} of {rendered['pages']}.", *(Image(data=base64.b64decode(data), format="jpeg") for data in images)]


@tool(READ_ONLY)
def get_binary_content(page_identifier: str, callback_id: str = "", file_name: str = "", save_to: str = "") -> dict[str, Any]:
    """Download an image (by callback_id) or an attached file (by its preferredName as file_name), both listed by
    get_page(include_objects=true). With save_to, write it to that file path instead of returning base64."""

    raw = _binary(page_identifier, callback_id, file_name)
    if save_to:
        target = Path(save_to).expanduser().resolve()
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_bytes(raw)
        return ok(path=str(target), bytes=len(raw))
    return ok(base64=base64.b64encode(raw).decode("ascii"), bytes=len(raw))


def _index_search(start_id: str, query: str) -> list[dict[str, Any]]:
    xml = onenote.call("find_pages", start_id=start_id, query=query, include_unindexed=False, display=False, schema=XML_SCHEMA_2013)["xml"]
    return of_type(parse_hierarchy(xml), "page")


@tool(READ_ONLY)
def search_pages(
    query: str = "",
    start_identifier: str = "",
    modified_after: str = "",
    modified_before: str = "",
    max_results: int = 20,
    include_snippets: bool = True,
    include_recycle_bin: bool = False,
    backend: str = "auto",
) -> dict[str, Any]:
    """Find pages containing every word of query, ranked (title matches first), within an optional notebook/section.

    modified_after / modified_before (ISO dates such as 2026-10-01) filter by last edit. With an empty query the
    result is simply the pages edited in that window, newest first.
    backend="auto" answers from the local text index (page text plus image OCR) and re-reads pages edited since they
    were indexed; while the index is still being built it asks OneNote's own index and falls back to a time-limited
    scan. backend="local_scan" always scans; backend="index" only asks OneNote.
    """

    terms = search.query_terms(query)
    after, before = search.parse_when(modified_after), search.parse_when(modified_before)
    if not terms and after is None and before is None:
        raise ValueError("Give a query, a modified_after/modified_before date, or both.")
    mode = backend.strip().casefold()
    if mode not in SEARCH_BACKENDS:
        raise ValueError(f"backend must be one of: {', '.join(sorted(SEARCH_BACKENDS))}")
    start_id = onenote.resolve(start_identifier)["id"] if start_identifier else ""
    limit = max(1, max_results)
    pages = [page for page in onenote.pages(start_id, include_recycle_bin) if search.in_window(page, after, before)]

    if not terms:
        pages.sort(key=lambda page: page.get("lastModifiedTime", ""), reverse=True)
        return ok(pages=pages[:limit], count=min(limit, len(pages)), total_matches=len(pages), search_backend="hierarchy")

    data: dict[str, Any] = {}
    stale = search.coverage(text_index, pages)
    if mode == "auto" and stale > search.STALE_PAGES_READ_INLINE:
        indexer.kick(force=True)
        try:
            allowed = {page["id"]: page for page in pages}
            hits = [allowed[hit["id"]] for hit in _index_search(start_id, query) if hit["id"] in allowed]
        except Exception as exc:
            hits, data["index_error"] = [], str(exc)
        if hits:
            data.update(pages=hits[:limit], search_backend="onenote_index", text_index=indexer.status())
    elif mode == "index":
        allowed = {page["id"]: page for page in pages}
        hits = [allowed[hit["id"]] for hit in _index_search(start_id, query) if hit["id"] in allowed]
        data.update(pages=hits[:limit], search_backend="onenote_index")
    if "pages" not in data:
        result = search.scan(onenote, text_index, pages, terms, limit=limit, budget_seconds=settings.scan_seconds)
        backend_used = "text_index" if mode == "auto" and stale <= search.STALE_PAGES_READ_INLINE else "local_scan"
        data.update(result, search_backend=backend_used)
        if not result["scan_complete"]:
            data["text_index"] = indexer.status()

    if include_snippets:
        for page in data["pages"]:
            try:
                text = search.page_text(onenote, text_index, page)
                page["snippet"] = search.snippet(text, terms) or text[:200]
            except Exception as exc:
                page["snippet_error"] = str(exc)
    return ok(count=len(data["pages"]), **data)


@tool(READ_ONLY)
def get_hyperlink(object_identifier: str, page_content_object_id: str = "", web: bool = False) -> dict[str, Any]:
    """A onenote: link (or, with web=true, a browser link) to a notebook, section, page or paragraph."""

    object_id = onenote.resolve(object_identifier)["id"]
    operation = "get_web_hyperlink" if web else "get_hyperlink"
    return ok(hyperlink=onenote.call(operation, object_id=object_id, page_content_object_id=page_content_object_id)["hyperlink"])
