"""Tools that create notebooks, sections and pages, and change page content."""

from __future__ import annotations

import base64
import copy
from pathlib import Path
from typing import Any

from ..app import DESTRUCTIVE, READ_ONLY, WRITES, backups, ok, onenote, tool
from ..constants import CREATE_FILE_TYPES, NEW_PAGE_STYLES, ONE, SPECIAL_LOCATIONS, XML_SCHEMA_2013, enum_value
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
def create_item(kind: str, name: str, parent_identifier: str = "") -> dict[str, Any]:
    """Create a notebook, section or section group (kind: notebook, section, section_group).

    A section or group goes under parent_identifier (a notebook or section group). A notebook is a folder: name is
    either a full folder path or a name placed in OneNote's default notebook folder.
    """

    kind = kind.strip().casefold()
    if kind == "notebook":
        raw = Path(name)
        if not raw.is_absolute():
            raw = Path(onenote.call("get_special_location", location=SPECIAL_LOCATIONS["default_notebook_folder"])["path"]) / safe_leaf_name(name)
        result = onenote.call("open_hierarchy", path=str(raw), relative_to_id="", create_file_type=CREATE_FILE_TYPES["notebook"])
        return ok(path=str(raw), notebook_id=result["object_id"], item=onenote.lookup(result["object_id"]))
    if kind not in {"section", "section_group"}:
        raise ValueError("kind must be notebook, section or section_group.")
    if not parent_identifier:
        raise ValueError("parent_identifier (a notebook or section group) is required for sections and section groups.")
    parent = onenote.resolve(parent_identifier)
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
def open_notebook(path: str) -> dict[str, Any]:
    """Open an existing notebook folder (or a .one section file) in OneNote. Opening one that is already open is fine."""

    result = onenote.call("open_hierarchy", path=path, relative_to_id="", create_file_type=CREATE_FILE_TYPES["none"])
    return ok(object_id=result["object_id"], item=onenote.lookup(result["object_id"]))


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


def _local_file(path: str) -> Path:
    file = Path(path).expanduser()
    if not file.is_file():
        raise ValueError(f"File not found: {path}")
    return file.resolve()


@tool(WRITES)
def append_to_page(
    page_identifier: str,
    content: str = "",
    content_format: str = "plain",
    image_path: str = "",
    file_path: str = "",
    display_name: str = "",
    x: float | None = None,
    y: float | None = None,
    width: float | None = None,
    height: float | None = None,
) -> dict[str, Any]:
    """Add one new block below the page's existing content (or at x/y): text, an image, or an attached file.

    Give exactly one of content (with content_format plain, html or markdown), image_path (a local image; width or
    height alone scales it proportionally) or file_path (any local file, shown as a file icon named display_name).
    To change existing text use edit_page.
    """

    given = [name for name, value in (("content", content), ("image_path", image_path), ("file_path", file_path)) if value]
    if len(given) != 1:
        raise ValueError("Give exactly one of content, image_path or file_path.")
    page_id = onenote.resolve(page_identifier, "page")["id"]
    if image_path:
        file = _local_file(image_path)
        fmt = file.suffix.lstrip(".")
        if not fmt:
            raise ValueError("image_path needs a file extension (png, jpg, gif, ...).")
        width, height = proportional_dimensions(file, width, height)
        x, y = below_existing(page_id, x, y)
        data = base64.b64encode(file.read_bytes()).decode("ascii")
        onenote.update_page(build_image_page_update_xml(page_id, image_base64=data, image_format=fmt, x=x, y=y, width=width, height=height))
        return ok(page_id=page_id, image_path=str(file), width=width, height=height)
    if file_path:
        file = _local_file(file_path)
        x, y = below_existing(page_id, x, y)
        onenote.update_page(build_file_page_update_xml(page_id, path=str(file), name=display_name or file.name, x=x, y=y))
        return ok(page_id=page_id, file=str(file), name=display_name or file.name)
    x, y = below_existing(page_id, x, y)
    onenote.update_page(build_page_update_xml(page_id, content=content, content_format=content_format, x=x, y=y))
    return ok(page_id=page_id, appended=True)


@tool(DESTRUCTIVE)
def edit_page(page_identifier: str, operations: list[EditOperation]) -> dict[str, Any]:
    """Change parts of a page without touching the rest (images, handwriting, attachments and other blocks stay as they are).

    Get paragraph IDs from get_page(include_paragraphs=true). op="delete" also takes the object_id of a whole
    block, image, ink drawing or attachment from get_page(include_objects=true). Operations run in order in one update, e.g.
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
    removed = [node.attrib["objectID"] for node in editor.removed]
    if any(node.tag != ONE + "Outline" for node in editor.removed):
        backup_id = snapshot(page_id, "edit_page")
    else:
        backup_id = snapshot_blocks(page_id, original, touched | set(removed), "edit_page")
    if update:
        onenote.update_page(update)
    for object_id in emptied + removed:
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
    title=None keeps the current title.

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
def submit_xml(xml: str, force: bool = False) -> dict[str, Any]:
    """Advanced: submit raw OneNote XML. A <one:Page> goes to UpdatePageContent (the page is backed up first);
    a <one:Notebooks>, <one:Notebook>, <one:SectionGroup> or <one:Section> goes to UpdateHierarchy."""

    root = parse_xml(xml)
    if root.tag != ONE + "Page":
        onenote.call("update_hierarchy", xml=xml, schema=XML_SCHEMA_2013)
        return ok(updated="hierarchy")
    page_id = root.attrib.get("ID", "")
    backup_id = snapshot(page_id, "submit_xml") if page_id else None
    onenote.update_page(xml, force=force)
    return ok(updated="page", page_id=page_id, backup_id=backup_id)


@tool(WRITES)
def copy_page(page_identifier: str, destination_section_identifier: str = "") -> dict[str, Any]:
    """Copy a page, with its images, handwriting and attachments, into a section (default: its own section)."""

    page = onenote.resolve(page_identifier, "page")
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
