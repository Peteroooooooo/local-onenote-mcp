"""Tools that create notebooks, sections and pages, and change page content."""

from __future__ import annotations

import base64
import copy
from pathlib import Path
from typing import Any

from ..app import DESTRUCTIVE, READ_ONLY, WRITES, backups, ok, onenote, tool
from ..constants import CREATE_FILE_TYPES, NEW_PAGE_STYLES, ONE, SPECIAL_LOCATIONS, enum_value
from ..editing import EditOperation, PageEditor
from ..hierarchy import ONENOTE_ID, child_path, in_recycle_bin, parse_xml, safe_leaf_name
from ..image_utils import proportional_dimensions
from ..pages import below_existing, replace_contents, restore_blocks, snapshot, snapshot_blocks, write_copy
from ..pagexml import (
    body_object_ids,
    build_file_page_update_xml,
    build_image_page_update_xml,
    build_page_update_xml,
    collect_page_objects,
)

RESTORE_MODES = {"auto", "new_page", "replace"}


@tool(WRITES)
def create_notebook(name_or_path: str, base_folder: str = "") -> dict[str, Any]:
    """Create a local notebook folder and open it in OneNote."""

    raw = Path(name_or_path)
    if not raw.is_absolute():
        root = base_folder or onenote.call("get_special_location", location=SPECIAL_LOCATIONS["default_notebook_folder"])["path"]
        raw = Path(root) / safe_leaf_name(name_or_path)
    result = onenote.call("open_hierarchy", path=str(raw), relative_to_id="", create_file_type=CREATE_FILE_TYPES["notebook"])
    return ok(path=str(raw), notebook_id=result["object_id"])


def _create_container(parent_identifier: str, name: str, kind: str) -> dict[str, Any]:
    parent = onenote.resolve(parent_identifier, full=True)
    if parent["type"] not in {"notebook", "section_group"}:
        raise ValueError("parent_identifier must resolve to a notebook or section_group.")
    leaf = safe_leaf_name(name)
    if kind == "section" and not leaf.lower().endswith(".one"):
        leaf += ".one"
    result = onenote.call("open_hierarchy", path=leaf, relative_to_id=parent["id"], create_file_type=CREATE_FILE_TYPES[kind])
    expected = child_path(parent["path"], leaf)
    item = onenote.wait_for(item_type=kind, path=expected, id=result["object_id"])
    return ok(parent=parent, item=item, path=expected, **{f"{kind}_id": item["id"] if item else result["object_id"]})


@tool(WRITES)
def create_section(parent_identifier: str, section_name: str) -> dict[str, Any]:
    """Create a section under a notebook or section group."""

    return _create_container(parent_identifier, section_name, "section")


@tool(WRITES)
def create_section_group(parent_identifier: str, group_name: str) -> dict[str, Any]:
    """Create a section group under a notebook or another section group."""

    return _create_container(parent_identifier, group_name, "section_group")


@tool(WRITES)
def open_hierarchy(path: str, relative_to_identifier: str = "", create_type: str = "none") -> dict[str, Any]:
    """Open a notebook folder, section file or section group, optionally creating it (create_type: none, notebook, section_group, section)."""

    kind = create_type.strip().casefold() or "none"
    relative_to = onenote.resolve(relative_to_identifier, full=True) if relative_to_identifier else None
    if kind == "none":
        expected = child_path(relative_to["path"], path) if relative_to else path.replace("\\", "/").strip("/")
        existing = onenote.find(path=expected)
        if existing:
            return ok(object_id=existing["id"], item=existing, opened_existing=True)
    result = onenote.call(
        "open_hierarchy",
        path=path,
        relative_to_id=relative_to["id"] if relative_to else "",
        create_file_type=enum_value("create_type", kind, CREATE_FILE_TYPES),
    )
    item = onenote.lookup(result["object_id"])
    return ok(object_id=result["object_id"], item=item, opened_existing=False)


@tool(WRITES)
def create_page(
    section_identifier: str,
    title: str,
    content: str = "",
    content_format: str = "plain",
    new_page_style: str = "blank_with_title",
) -> dict[str, Any]:
    """Create a page in a section. content_format: plain, html or markdown."""

    section = onenote.resolve(section_identifier, "section")
    xml = build_page_update_xml("pending", title=title, content=content, content_format=content_format)
    page_id = onenote.new_page(section["id"], enum_value("new_page_style", new_page_style, NEW_PAGE_STYLES))
    onenote.update_page(xml.replace('ID="pending"', f'ID="{page_id}"', 1))
    return ok(page_id=page_id, page=onenote.wait_for(item_type="page", id=page_id), title=title)


@tool(WRITES)
def update_page_title(page_identifier: str, title: str) -> dict[str, Any]:
    """Change a page's title."""

    page_id = onenote.resolve(page_identifier, "page")["id"]
    onenote.update_page(build_page_update_xml(page_id, title=title))
    return ok(page_id=page_id, title=title)


@tool(WRITES)
def append_to_page(
    page_identifier: str,
    content: str,
    content_format: str = "plain",
    x: float | None = None,
    y: float | None = None,
) -> dict[str, Any]:
    """Add a new block below the page's existing content (or at x/y). content_format: plain, html, markdown."""

    page_id = onenote.resolve(page_identifier, "page")["id"]
    x, y = below_existing(page_id, x, y)
    onenote.update_page(build_page_update_xml(page_id, content=content, content_format=content_format, x=x, y=y))
    return ok(page_id=page_id, appended=True)


def _local_file(path: str) -> Path:
    file = Path(path).expanduser()
    if not file.is_file():
        raise ValueError(f"File not found: {path}")
    return file.resolve()


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
    """Add a local image below the page's content (or at x/y). Give width or height alone to scale proportionally."""

    file = _local_file(image_path)
    fmt = image_format or file.suffix.lstrip(".")
    if not fmt:
        raise ValueError("image_format is required when image_path has no extension.")
    width, height = proportional_dimensions(file, width, height)
    page_id = onenote.resolve(page_identifier, "page")["id"]
    x, y = below_existing(page_id, x, y)
    data = base64.b64encode(file.read_bytes()).decode("ascii")
    onenote.update_page(build_image_page_update_xml(page_id, image_base64=data, image_format=fmt, x=x, y=y, width=width, height=height))
    return ok(page_id=page_id, image_path=str(file), width=width, height=height)


@tool(WRITES)
def add_file_to_page(
    page_identifier: str,
    file_path: str,
    display_name: str = "",
    x: float | None = None,
    y: float | None = None,
) -> dict[str, Any]:
    """Attach a local file (PDF, document, archive, ...) to a page as a file icon, below the existing content unless x/y are given."""

    file = _local_file(file_path)
    page_id = onenote.resolve(page_identifier, "page")["id"]
    x, y = below_existing(page_id, x, y)
    onenote.update_page(build_file_page_update_xml(page_id, path=str(file), name=display_name or file.name, x=x, y=y))
    return ok(page_id=page_id, file=str(file), name=display_name or file.name)


@tool(DESTRUCTIVE)
def edit_page(page_identifier: str, operations: list[EditOperation]) -> dict[str, Any]:
    """Change parts of a page without touching the rest (images, handwriting, attachments and other blocks stay as they are).

    Get paragraph IDs from get_page(include_paragraphs=true). Operations run in order in one update, e.g.
    [{"op": "replace_text", "find": "2025", "replace": "2026"},
     {"op": "insert_after", "paragraph_id": "...", "content": "- new item"},
     {"op": "set_todo", "paragraph_id": "...", "done": true}].
    A backup of the changed blocks is saved first. Returns the page's paragraphs after the edit.
    """

    if not operations:
        raise ValueError("operations is empty.")
    ops = [EditOperation.model_validate(op) if isinstance(op, dict) else op for op in operations]
    page_id = onenote.resolve(page_identifier, "page")["id"]
    page = parse_xml(onenote.page_xml(page_id))
    original = copy.deepcopy(page)
    editor = PageEditor(page)
    results = [editor.apply(op) for op in ops]
    update, emptied = editor.build_update(page_id), editor.empty_outlines()
    touched = {outline.attrib["objectID"] for outline in editor.touched if outline.attrib.get("objectID")}
    backup_id = snapshot_blocks(page_id, original, touched, "edit_page")
    if update:
        onenote.update_page(update)
    for object_id in emptied:
        onenote.delete_object(page_id, object_id)
    paragraphs = PageEditor(parse_xml(onenote.page_xml(page_id))).paragraphs()
    return ok(page_id=page_id, results=results, backup_id=backup_id, paragraphs=paragraphs)


@tool(DESTRUCTIVE)
def replace_page_body(
    page_identifier: str,
    content: str,
    title: str | None = None,
    content_format: str = "plain",
) -> dict[str, Any]:
    """Replace everything on a page (text, images, ink, attachments) with new content. Prefer edit_page for partial changes.

    The new content is written before the old objects are deleted, and a backup is saved first.
    """

    page_id = onenote.resolve(page_identifier, "page")["id"]
    old_ids = body_object_ids(parse_xml(onenote.page_xml(page_id)))
    backup_id = snapshot(page_id, "replace_page_body")
    onenote.update_page(build_page_update_xml(page_id, title=title, content=content, content_format=content_format), force=True)
    for object_id in old_ids:
        onenote.delete_object(page_id, object_id)
    return ok(page_id=page_id, deleted_objects=old_ids, backup_id=backup_id)


@tool(DESTRUCTIVE)
def delete_page_content(page_identifier: str, object_id: str) -> dict[str, Any]:
    """Delete one outline, image, ink drawing or attachment from a page. Get object IDs from get_page(include_objects=true).

    To delete single paragraphs use edit_page. A backup is saved first.
    """

    page_id = onenote.resolve(page_identifier, "page")["id"]
    match = next((obj for obj in collect_page_objects(onenote.page_xml(page_id)) if obj.get("object_id") == object_id), None)
    if match is None:
        raise ValueError(f"Object '{object_id}' is not on this page.")
    if not match["delete_supported"]:
        hint = f" Delete its block '{match['delete_object_id']}', or use edit_page op=delete." if match.get("delete_object_id") else ""
        raise ValueError(f"'{object_id}' is a {match['type']} inside a block and cannot be deleted on its own.{hint}")
    backup_id = snapshot(page_id, "delete_page_content")
    onenote.delete_object(page_id, object_id)
    return ok(page_id=page_id, object_id=object_id, backup_id=backup_id)


@tool(DESTRUCTIVE)
def update_page_xml(xml: str, force: bool = False) -> dict[str, Any]:
    """Advanced: submit raw OneNote page XML to UpdatePageContent. A backup of the page is saved first."""

    page_id = parse_xml(xml).attrib.get("ID", "")
    backup_id = snapshot(page_id, "update_page_xml") if page_id else None
    onenote.update_page(xml, force=force)
    return ok(page_id=page_id, backup_id=backup_id)


@tool(WRITES)
def copy_page(page_identifier: str, destination_section_identifier: str = "") -> dict[str, Any]:
    """Copy a page, with its images, handwriting and attachments, into a section (default: its own section)."""

    page = onenote.resolve(page_identifier, "page", full=True)
    section_id = onenote.resolve(destination_section_identifier, "section")["id"] if destination_section_identifier else page["parent_id"]
    new_id = write_copy(parse_xml(onenote.page_xml(page["id"], "all")), section_id)
    return ok(page_id=new_id, page=onenote.wait_for(item_type="page", id=new_id), copied_from=page["id"])


@tool(READ_ONLY)
def list_backups(page_identifier: str = "", limit: int = 20) -> dict[str, Any]:
    """List saved page backups, newest first; optionally only those of one page (an ID also works for deleted pages)."""

    page_id = ""
    if page_identifier:
        page_id = page_identifier if ONENOTE_ID.match(page_identifier.strip()) else onenote.resolve(page_identifier, "page")["id"]
    return ok(backups=backups.list(page_id.strip(), limit), storage=backups.stats())


@tool(WRITES)
def restore_backup(backup_id: int, mode: str = "auto", section_identifier: str = "") -> dict[str, Any]:
    """Bring back a page, or the blocks an edit_page changed, from list_backups.

    mode="replace" puts the saved version back into the original page; its current state is backed up first,
    so a restore can itself be undone. mode="new_page" writes a saved whole page as a new page in
    section_identifier, or in the page's original section (this also works after the page was deleted).
    mode="auto" (default): "replace" for backups of kind "blocks", "new_page" for whole pages.
    """

    mode = mode.strip().casefold()
    if mode not in RESTORE_MODES:
        raise ValueError("mode must be 'auto', 'new_page' or 'replace'.")
    meta, page, files = backups.load(backup_id)
    blocks = meta["kind"] == "blocks"
    if mode == "auto":
        mode = "replace" if blocks else "new_page"
    if mode == "replace":
        target = onenote.lookup(meta["page_id"])
        if target is None or in_recycle_bin(target):
            hint = " A blocks backup cannot be restored without its page." if blocks else " Use mode='new_page'."
            raise ValueError("The original page no longer exists." + hint)
        if blocks:
            current = parse_xml(onenote.page_xml(meta["page_id"]))
            ids = {node.attrib.get("objectID") for node in page.findall(ONE + "Outline")} - {None}
            before = snapshot_blocks(meta["page_id"], current, ids, f"before restore_backup {backup_id}")
            restore_blocks(meta["page_id"], page)
        else:
            before = snapshot(meta["page_id"], f"before restore_backup {backup_id}")
            replace_contents(meta["page_id"], page, files)
        return ok(page_id=meta["page_id"], restored_from=backup_id, backup_id=before)
    if blocks:
        raise ValueError("This backup holds only the blocks an edit changed; restore it with mode='replace'.")
    if section_identifier:
        section_id = onenote.resolve(section_identifier, "section")["id"]
    else:
        section = onenote.lookup(meta["section_id"]) if meta["section_id"] else None
        if section is None or in_recycle_bin(section):
            raise ValueError("The page's original section no longer exists; pass section_identifier.")
        section_id = section["id"]
    new_id = write_copy(page, section_id, files)
    return ok(page_id=new_id, page=onenote.wait_for(item_type="page", id=new_id), restored_from=backup_id)
