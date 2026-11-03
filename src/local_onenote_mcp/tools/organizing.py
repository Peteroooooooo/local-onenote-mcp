"""Tools that reorganise the hierarchy: move, indent, rename, merge and delete."""

from __future__ import annotations

import time
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Any

from ..app import DESTRUCTIVE, WRITES, ok, onenote, tool
from ..constants import ONE, XML_SCHEMA_2013
from ..hierarchy import Item, in_recycle_bin, of_type, parse_hierarchy, safe_leaf_name
from ..pages import snapshot

FILE_MOVE_WAIT_SECONDS = 20.0


def _pages(section: ET.Element) -> list[ET.Element]:
    return [node for node in section if node.tag == ONE + "Page"]


def _level(node: ET.Element) -> int:
    try:
        return int(node.attrib.get("pageLevel", "1"))
    except ValueError:
        return 1


def _block(pages: list[ET.Element], index: int, include_subpages: bool = True) -> list[ET.Element]:
    """The page at index plus, optionally, the subpages directly below it."""

    block = [pages[index]]
    if include_subpages:
        for node in pages[index + 1 :]:
            if _level(node) <= _level(pages[index]):
                break
            block.append(node)
    return block


def _index_of(pages: list[ET.Element], page_id: str) -> int:
    for index, node in enumerate(pages):
        if node.attrib.get("ID") == page_id:
            return index
    raise RuntimeError("The page was not found in its section; refresh and try again.")


def _settle(notebook_ids: set[str], old_file: str) -> bool:
    """Sync, then wait until OneNote has moved a local section's .one file; otherwise it can reappear on close."""

    for notebook_id in notebook_ids:
        onenote.call("sync_hierarchy", object_id=notebook_id)
    if not old_file or "://" in old_file:
        return True
    deadline = time.monotonic() + FILE_MOVE_WAIT_SECONDS
    while Path(old_file).exists() and time.monotonic() < deadline:
        time.sleep(0.25)
    return not Path(old_file).exists()


def _live(identifier: str, item_type: str | None = None) -> Item:
    item = onenote.resolve(identifier, item_type, full=True)
    if in_recycle_bin(item):
        raise ValueError(f"'{item['path']}' is in the recycle bin; restore it in OneNote first.")
    return item


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
    include_subpages=false, and it becomes a top-level page (set_page_level indents it).
    Moving to another section gives the pages new IDs; the old-to-new mapping is returned.
    """

    page = _live(page_identifier, "page")
    source_id = page["parent_id"]
    destination_id = _live(destination_section_identifier, "section")["id"] if destination_section_identifier else source_id
    same_section = destination_id == source_id

    source = onenote.tree(source_id, "pages")
    block = _block(_pages(source), _index_of(_pages(source), page["id"]), include_subpages)
    block_ids = [node.attrib["ID"] for node in block]
    target = source if same_section else onenote.tree(destination_id, "pages")
    if same_section:
        for node in block:
            target.remove(node)
    known_ids = {node.attrib.get("ID") for node in _pages(target)}

    shift = _level(block[0]) - 1
    for node in block:
        node.set("pageLevel", str(max(1, _level(node) - shift)))

    pages, children = _pages(target), list(target)
    if after_page_identifier:
        after_id = onenote.resolve(after_page_identifier, "page")["id"]
        if after_id in block_ids:
            raise ValueError("after_page_identifier cannot be the page being moved or one of its subpages.")
        if after_id not in known_ids:
            raise ValueError("after_page_identifier must be a page in the destination section.")
        insert_at = children.index(_block(pages, _index_of(pages, after_id))[-1]) + 1
    elif to_top and pages:
        insert_at = children.index(pages[0])
    else:
        insert_at = children.index(pages[-1]) + 1 if pages else len(children)
    for offset, node in enumerate(block):
        target.insert(insert_at + offset, node)
    onenote.submit_tree(target)

    if same_section:
        new_ids = block_ids
    else:
        for _ in range(20):
            new_ids = [node.attrib["ID"] for node in _pages(onenote.tree(destination_id, "pages")) if node.attrib.get("ID") not in known_ids]
            if len(new_ids) >= len(block):
                break
            time.sleep(0.25)
        left = {node.attrib.get("ID") for node in _pages(onenote.tree(source_id, "pages"))} & set(block_ids)
        if len(new_ids) < len(block) or left:
            raise RuntimeError(
                f"OneNote accepted the move but it could not be confirmed ({len(new_ids)} of {len(block)} pages arrived, "
                f"{len(left)} still in the source)."
            )
    moved = [{"old_id": old, "new_id": new} for old, new in zip(block_ids, new_ids)]
    return ok(page_id=new_ids[0], page=onenote.lookup(new_ids[0]), moved_pages=moved, destination_section_id=destination_id)


@tool(WRITES)
def set_page_level(page_identifier: str, level: int) -> dict[str, Any]:
    """Indent a page as a subpage (level 2 or 3) or promote it back to a top-level page (level 1)."""

    if level not in {1, 2, 3}:
        raise ValueError("level must be 1, 2 or 3.")
    page = onenote.resolve(page_identifier, "page", full=True)
    tree = onenote.tree(page["parent_id"], "pages")
    pages = _pages(tree)
    index = _index_of(pages, page["id"])
    if level > 1 and (index == 0 or _level(pages[index - 1]) < level - 1):
        raise ValueError("A subpage needs a page directly above it one level higher; move the page first.")
    pages[index].set("pageLevel", str(level))
    onenote.submit_tree(tree)
    return ok(page_id=page["id"], level=level)


def _siblings(tree: ET.Element, tag: str) -> list[ET.Element]:
    return [node for node in tree if node.tag == ONE + tag and node.attrib.get("isRecycleBin") != "true"]


def _container(identifier: str) -> Item:
    item = _live(identifier)
    if item["type"] not in {"section", "section_group"}:
        raise ValueError("The identifier must resolve to a section or section_group.")
    return item


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

    item = _container(section_identifier)
    destination = _live(destination_identifier) if destination_identifier else onenote.resolve(item["parent_id"], full=True)
    if destination["type"] not in {"notebook", "section_group"}:
        raise ValueError("destination_identifier must resolve to a notebook or section_group.")
    if destination["id"] == item["id"] or destination["path"].startswith(item["path"] + "/"):
        raise ValueError("A section group cannot be moved into itself.")

    tag = "Section" if item["type"] == "section" else "SectionGroup"
    same_parent = destination["id"] == item["parent_id"]
    target = onenote.tree(destination["id"], "sections")
    if same_parent:
        node = next(child for child in target if child.attrib.get("ID") == item["id"])
        target.remove(node)
    else:
        if any(child.attrib.get("name", "").casefold() == item["name"].casefold() for child in _siblings(target, tag)):
            raise ValueError(f"'{destination['path']}' already has a {item['type']} named '{item['name']}'; rename one first.")
        node = onenote.tree(item["id"], "sections")

    siblings, children = _siblings(target, tag), list(target)
    if after_identifier:
        after_id = onenote.resolve(after_identifier)["id"]
        anchor = next((child for child in siblings if child.attrib.get("ID") == after_id), None)
        if anchor is None:
            raise ValueError(f"after_identifier must be a {item['type']} directly inside '{destination['path']}'.")
        insert_at = children.index(anchor) + 1
    elif siblings:
        insert_at = children.index(siblings[0]) if to_top else children.index(siblings[-1]) + 1
    else:
        # Sections come before section groups, and the recycle bin stays last.
        later = [child for child in children if child.tag == ONE + "SectionGroup" and (tag == "Section" or child.attrib.get("isRecycleBin") == "true")]
        insert_at = children.index(later[0]) if later else len(children)
    target.insert(insert_at, node)
    onenote.submit_tree(target)

    settled = same_parent or _settle({onenote.notebook_of(item), onenote.notebook_of(destination)}, item.get("onenote_path", ""))
    moved = onenote.find(id=item["id"], item_type=item["type"], fresh=True)
    if moved is None or not moved["path"].startswith(destination["path"] + "/"):
        raise RuntimeError("OneNote accepted the move but the item is not under the destination yet; check OneNote.")
    data = ok(object_id=item["id"], item=moved, destination=destination["path"])
    if not settled:
        data["warning"] = "OneNote has not finished moving the section file on disk; keep OneNote open a moment before closing the notebook."
    return data


@tool(WRITES)
def rename_section(section_identifier: str, new_name: str) -> dict[str, Any]:
    """Rename a section or section group. Rename pages with update_page_title."""

    item = _container(section_identifier)
    name = safe_leaf_name(new_name)
    tag = "Section" if item["type"] == "section" else "SectionGroup"
    parent = onenote.tree(item["parent_id"], "sections")
    for child in _siblings(parent, tag):
        if child.attrib.get("ID") != item["id"] and child.attrib.get("name", "").casefold() == name.casefold():
            raise ValueError(f"A {item['type']} named '{name}' already exists next to it.")
    next(child for child in parent if child.attrib.get("ID") == item["id"]).set("name", name)
    onenote.submit_tree(parent)
    settled = _settle({onenote.notebook_of(item)}, item.get("onenote_path", ""))
    data = ok(object_id=item["id"], item=onenote.find(id=item["id"], item_type=item["type"], fresh=True), name=name)
    if not settled:
        data["warning"] = "OneNote has not finished renaming the section file on disk yet."
    return data


@tool(DESTRUCTIVE)
def merge_sections(source_section_identifier: str, destination_section_identifier: str) -> dict[str, Any]:
    """Move every page of one section into another (OneNote's Merge into Another Section)."""

    source_id = onenote.resolve(source_section_identifier, "section")["id"]
    destination_id = onenote.resolve(destination_section_identifier, "section")["id"]
    onenote.call("merge_sections", source_section_id=source_id, destination_section_id=destination_id)
    return ok(source_section_id=source_id, destination_section_id=destination_id, merged=True)


def _page_ids_under(item: Item) -> list[str]:
    if item["type"] == "page":
        return [item["id"]]
    return [page["id"] for page in of_type(parse_hierarchy(onenote.hierarchy_xml(item["id"], "pages")), "page")]


@tool(DESTRUCTIVE)
def delete_hierarchy(object_identifier: str, permanently: bool = False) -> dict[str, Any]:
    """Delete a notebook's section group, a section or a page. permanently=false moves it to OneNote's recycle bin.

    Pages are backed up before a delete; a permanently deleted section has each of its pages backed up.
    """

    item = onenote.resolve(object_identifier, full=True)
    if item["type"] == "notebook":
        raise ValueError("Notebooks are folders on disk; close them with close_notebook and delete the folder yourself.")
    backed_up = []
    if item["type"] == "page" or permanently:
        backed_up = [backup for page_id in _page_ids_under(item) if (backup := snapshot(page_id, "delete_hierarchy")) is not None]

    if item["type"] == "page":
        onenote.call("delete_hierarchy", object_id=item["id"], permanently=permanently)
        time.sleep(0.5)
        remaining = onenote.find(id=item["id"], item_type="page", fresh=True)
        if remaining and remaining["path"] == item["path"]:
            raise RuntimeError(f"Delete returned success, but page '{item['path']}' is still there.")
        return ok(deleted_ids=[item["id"]], permanently=permanently, backup_ids=backed_up)

    # OneNote can re-create a section or group under a new ID right after deleting
    # it; their names are unique within a folder, so retry by path.
    deleted = []
    for _ in range(4):
        onenote.call("delete_hierarchy", object_id=item["id"], permanently=permanently)
        deleted.append(item["id"])
        time.sleep(0.5)
        remaining = onenote.find(path=item["path"], item_type=item["type"], fresh=True)
        if remaining is None:
            return ok(deleted_ids=deleted, permanently=permanently, backup_ids=backed_up)
        item = remaining
    raise RuntimeError(f"Delete returned success, but '{item['path']}' still exists with ID {item['id']}.")


@tool(DESTRUCTIVE)
def update_hierarchy_xml(xml: str) -> dict[str, Any]:
    """Advanced: submit raw OneNote hierarchy XML to UpdateHierarchy."""

    onenote.call("update_hierarchy", xml=xml, schema=XML_SCHEMA_2013)
    return ok(updated=True)
