"""Tools that reorganise the hierarchy: move, indent, rename, merge and delete."""

from __future__ import annotations

import time
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Any

from ..app import DESTRUCTIVE, WRITES, ok, onenote, tool
from ..constants import CREATE_FILE_TYPES
from ..constants import ONE
from ..hierarchy import Item, in_recycle_bin, of_type, parse_hierarchy, safe_leaf_name
from ..pages import snapshot
from ..pagexml import build_page_update_xml

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


def _submit_pages(section: ET.Element, before: dict[str, int], moved: set[str]) -> list[dict[str, Any]]:
    """Submit a section's page order, first making every page at most one level deeper than the page above it.

    OneNote crashes (onmain.dll access violation) on a page list with a level jump such as 1 -> 3, and it can
    leave such jumps behind itself, e.g. after the parent of a level-3 page was deleted.
    Returns the pages outside `moved` whose level changed, so the caller can tell the user.
    """

    previous = 0
    for node in _pages(section):
        level = min(3, max(1, min(_level(node), previous + 1)))
        if level != _level(node):
            node.set("pageLevel", str(level))
        previous = level
    onenote.submit_tree(section)
    return [
        {"page": node.attrib.get("name", ""), "page_id": node.attrib["ID"], "from_level": before[node.attrib["ID"]], "to_level": _level(node)}
        for node in _pages(section)
        if node.attrib.get("ID") in before and node.attrib["ID"] not in moved and before[node.attrib["ID"]] != _level(node)
    ]


def _levels(section: ET.Element) -> dict[str, int]:
    return {node.attrib["ID"]: _level(node) for node in _pages(section) if node.attrib.get("ID")}


def _shift_block(block: list[ET.Element], new_level: int) -> None:
    """Give the block's first page new_level, keeping its subpages' depth below it; refuse to go past level 3."""

    shift = new_level - _level(block[0])
    deepest = max(_level(node) for node in block) + shift
    if deepest > 3:
        raise ValueError(
            f"That would put subpages at level {deepest}; OneNote has only 3 levels. "
            "Use a smaller level, include_subpages=false, or move the deepest subpages first."
        )
    for node in block:
        node.set("pageLevel", str(max(1, _level(node) + shift)))


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
    return onenote.live(identifier, item_type)


@tool(WRITES)
def move_page(
    page_identifier: str,
    destination_section_identifier: str = "",
    after_page_identifier: str = "",
    to_top: bool = False,
    include_subpages: bool = True,
    level: int = 0,
) -> dict[str, Any]:
    """Move a page to another section (any notebook) or to a new position in its own section, and/or indent it.

    Leave destination_section_identifier empty to reorder within the current section. The page goes after
    after_page_identifier, to the top with to_top=true, or to the end. Its subpages move with it unless
    include_subpages=false. A moved page becomes a top-level page unless level is given; placed after a subpage
    that has sibling subpages below it, it joins that group at the subpage's level.
    level=2 or 3 makes the page a subpage of the page above it, level=1 promotes it; with only level given the
    page stays where it is.
    Moving to another section gives the pages new IDs; the old-to-new mapping is returned.
    """

    if level not in {0, 1, 2, 3}:
        raise ValueError("level must be 1, 2 or 3 (0 = leave as moved).")
    if level and not (destination_section_identifier or after_page_identifier or to_top):
        page = _live(page_identifier, "page")
        others = _set_level(page["id"], page["parent_id"], level, include_subpages)
        return _with_side_effects(ok(page_id=page["id"], level=level), others)
    if to_top and level > 1:
        raise ValueError("The first page of a section cannot be a subpage; use level=1 or another position.")
    page = _live(page_identifier, "page")
    source_id = page["parent_id"]
    destination_id = _live(destination_section_identifier, "section")["id"] if destination_section_identifier else source_id
    same_section = destination_id == source_id

    source = onenote.tree(source_id, "pages")
    index = _index_of(_pages(source), page["id"])
    block = _block(_pages(source), index, include_subpages)
    block_ids = [node.attrib["ID"] for node in block]
    # Subpages left behind rise one level, so they do not end up under whichever page was above their parent.
    left_behind = [] if include_subpages else _block(_pages(source), index)[1:]
    source_before = _levels(source)
    for node in left_behind:
        node.set("pageLevel", str(max(1, _level(node) - 1)))
    promoted: list[dict[str, Any]] = []
    if left_behind and not same_section:
        promoted = _submit_pages(source, source_before, set(block_ids))
        source = onenote.tree(source_id, "pages")
    target = source if same_section else onenote.tree(destination_id, "pages")
    before = source_before if same_section else _levels(target)
    if same_section:
        for node in block:
            target.remove(node)
    known_ids = {node.attrib.get("ID") for node in _pages(target)}

    pages, children = _pages(target), list(target)
    above: ET.Element | None = None
    if after_page_identifier:
        after_id = onenote.resolve(after_page_identifier, "page")["id"]
        if after_id in block_ids:
            raise ValueError("after_page_identifier cannot be the page being moved or one of its subpages.")
        if after_id not in known_ids:
            raise ValueError("after_page_identifier must be a page in the destination section.")
        after_block = _block(pages, _index_of(pages, after_id))
        above = after_block[-1]
        insert_at = children.index(above) + 1
        below = pages[pages.index(above) + 1] if pages.index(above) + 1 < len(pages) else None
        new_level = level or 1
        if below is not None and _level(below) > new_level:
            # A page dropped in front of deeper pages would adopt them as its subpages, and OneNote
            # crashes on that hierarchy update; take the level of the page it follows instead.
            if level:
                raise ValueError(
                    f"level={level} would put the page in the middle of a group of level-{_level(below)} subpages; "
                    f"use level={_level(after_block[0])} or move it after the group's last page."
                )
            new_level = _level(after_block[0])
    elif to_top and pages:
        insert_at = children.index(pages[0])
        new_level = 1
    else:
        insert_at = children.index(pages[-1]) + 1 if pages else len(children)
        above = pages[-1] if pages else None
        new_level = level or 1
    if new_level > 1 and (above is None or _level(above) < new_level - 1):
        raise ValueError("A subpage needs a page directly above it one level higher; choose another position or level.")

    _shift_block(block, new_level)
    for offset, node in enumerate(block):
        target.insert(insert_at + offset, node)
    others = _submit_pages(target, before, set(block_ids))

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
    data = ok(page_id=new_ids[0], page=onenote.lookup(new_ids[0]), moved_pages=moved, destination_section_id=destination_id)
    return _with_side_effects(data, promoted + others)


def _with_side_effects(data: dict[str, Any], others: list[dict[str, Any]]) -> dict[str, Any]:
    if others:
        data["other_level_changes"] = others
        data["note"] = (
            "OneNote nests pages by position: other pages changed level with this move (listed in other_level_changes)."
        )
    return data


def _set_level(page_id: str, section_id: str, level: int, include_subpages: bool = True) -> list[dict[str, Any]]:
    """Indent a page (its subpages keep their depth below it, unless include_subpages=false) or promote it."""

    tree = onenote.tree(section_id, "pages")
    before = _levels(tree)
    pages = _pages(tree)
    index = _index_of(pages, page_id)
    if level > 1 and (index == 0 or _level(pages[index - 1]) < level - 1):
        raise ValueError("A subpage needs a page directly above it one level higher; move the page first.")
    block = _block(pages, index, include_subpages)
    _shift_block(block, level)
    return _submit_pages(tree, before, {node.attrib["ID"] for node in block})


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
    after_identifier (a sibling of the same type), to the top with to_top=true, or to the end. IDs are kept,
    except that a section group moved to another notebook is recreated there (new group ID, same section IDs).
    For local notebooks the tool waits until OneNote has moved the .one file on disk.
    """

    return _move_section(section_identifier, destination_identifier, after_identifier, to_top)


def _move_section(section_identifier: str, destination_identifier: str, after_identifier: str, to_top: bool) -> dict[str, Any]:
    item = _container(section_identifier)
    destination = _live(destination_identifier) if destination_identifier else onenote.resolve(item["parent_id"])
    if destination["type"] not in {"notebook", "section_group"}:
        raise ValueError("destination_identifier must resolve to a notebook or section_group.")
    if destination["id"] == item["id"] or destination["path"].startswith(item["path"] + "/"):
        raise ValueError("A section group cannot be moved into itself.")

    tag = "Section" if item["type"] == "section" else "SectionGroup"
    same_parent = destination["id"] == item["parent_id"]
    if tag == "SectionGroup" and (after_identifier or to_top or same_parent):
        raise ValueError(
            "OneNote keeps section groups sorted by name, so their position cannot be chosen; "
            "move it into another notebook or group, or rename it to change where it sorts."
        )
    if after_identifier and onenote.resolve(after_identifier)["id"] == item["id"]:
        raise ValueError("after_identifier is the item being moved; give the sibling it should follow.")
    if tag == "SectionGroup" and not same_parent and onenote.notebook_of(item) != onenote.notebook_of(destination):
        return _move_group_to_notebook(item, destination)
    target = onenote.tree(destination["id"], "sections")
    if same_parent:
        node = next(child for child in target if child.attrib.get("ID") == item["id"])
        target.remove(node)
    else:
        clash = next((child for child in _siblings(target, tag) if child.attrib.get("name", "").casefold() == item["name"].casefold()), None)
        if clash is not None:
            raise ValueError(f"'{destination['path']}' already has a {item['type']} named '{clash.attrib.get('name')}'; rename one first.")
        node = onenote.tree(item["id"], "sections")
        # Moving up into an ancestor: its tree still holds the item further down; submitting both copies
        # leaves the item where it was.
        for parent in list(target.iter()):
            for child in list(parent):
                if child.attrib.get("ID") == item["id"]:
                    parent.remove(child)

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
    moved = _wait_for_parent(item, destination["id"])
    if moved is None:
        raise RuntimeError("OneNote accepted the move but the item is not under the destination yet; check OneNote.")
    data = ok(object_id=item["id"], item=moved, destination=destination["path"])
    if not settled:
        data["warning"] = "OneNote has not finished moving the section file on disk; keep OneNote open a moment before closing the notebook."
    return data


def _wait_for_parent(item: Item, parent_id: str, tries: int = 8) -> Item | None:
    for attempt in range(tries):
        found = onenote.find(id=item["id"], item_type=item["type"], fresh=True)
        if found is not None and found["parent_id"] == parent_id:
            return found
        if attempt + 1 < tries:
            time.sleep(0.25)
    return None


def _move_group_to_notebook(item: Item, destination: Item) -> dict[str, Any]:
    """Move a section group into another notebook by recreating it there and moving its contents into it.

    OneNote crashes or leaves a copy behind when a section group itself is moved across notebooks; sections
    move across notebooks fine. The sections keep their IDs; the group (and any groups inside it) get new ones.
    """

    target = onenote.tree(destination["id"], "sections")
    clash = next((child for child in _siblings(target, "SectionGroup") if child.attrib.get("name", "").casefold() == item["name"].casefold()), None)
    if clash is not None:
        raise ValueError(f"'{destination['path']}' already has a section_group named '{clash.attrib.get('name')}'; rename one first.")
    result = onenote.call("open_hierarchy", path=item["name"], relative_to_id=destination["id"], create_file_type=CREATE_FILE_TYPES["section_group"])
    group = onenote.wait_for(item_type="section_group", path=f"{destination['path']}/{item['name']}", id=result["object_id"])
    if group is None:
        raise RuntimeError(f"Could not create section group '{item['name']}' in '{destination['path']}'.")
    source = onenote.tree(item["id"], "sections")
    moved = []
    for child in [node for node in source if node.attrib.get("isRecycleBin") != "true" and node.tag in {ONE + "Section", ONE + "SectionGroup"}]:
        moved.append(_move_section(child.attrib["ID"], group["id"], "", False)["item"]["path"])
    left = [node for node in onenote.tree(item["id"], "sections") if node.tag in {ONE + "Section", ONE + "SectionGroup"}]
    if left:
        raise RuntimeError(f"Moved {len(moved)} items, but {len(left)} are still in '{item['path']}'; nothing was deleted.")
    onenote.call("delete_hierarchy", object_id=item["id"], permanently=False)
    return ok(
        object_id=group["id"],
        old_object_id=item["id"],
        item=onenote.find(id=group["id"], item_type="section_group", fresh=True) or group,
        destination=destination["path"],
        moved=moved,
        note="Moved across notebooks by recreating the section group; it has a new ID, its sections kept theirs.",
    )


@tool(WRITES)
def rename_item(identifier: str, new_name: str) -> dict[str, Any]:
    """Rename a page (its title), a section or a section group."""

    item = _live(identifier)
    if item["type"] == "page":
        title = " ".join(new_name.split())
        if not title:
            raise ValueError("A page title cannot be empty.")
        onenote.update_page(build_page_update_xml(item["id"], title=title))
        return ok(object_id=item["id"], name=title)
    if item["type"] not in {"section", "section_group"}:
        raise ValueError("Only pages, sections and section groups can be renamed here.")
    name = safe_leaf_name(new_name)
    if name == item["name"]:
        return ok(object_id=item["id"], item=item, name=name, unchanged=True)
    tag = "Section" if item["type"] == "section" else "SectionGroup"
    parent = onenote.tree(item["parent_id"], "sections")
    for child in _siblings(parent, tag):
        if child.attrib.get("ID") != item["id"] and child.attrib.get("name", "").casefold() == name.casefold():
            raise ValueError(f"A {item['type']} named '{name}' already exists next to it.")
    next(child for child in parent if child.attrib.get("ID") == item["id"]).set("name", name)
    onenote.submit_tree(parent)
    # A case-only rename keeps the same file name as far as Windows is concerned; there is no old file to wait for.
    old_file = "" if name.casefold() == item["name"].casefold() else item.get("onenote_path", "")
    settled = _settle({onenote.notebook_of(item)}, old_file)
    data = ok(object_id=item["id"], item=onenote.find(id=item["id"], item_type=item["type"], fresh=True), name=name)
    if not settled:
        data["warning"] = "OneNote has not finished renaming the section file on disk yet."
    return data


@tool(DESTRUCTIVE)
def merge_sections(source_section_identifier: str, destination_section_identifier: str) -> dict[str, Any]:
    """Move every page of one section to the end of another (OneNote's Merge into Another Section).

    The emptied source section then goes to the recycle bin, as OneNote's own merge offers.
    """

    source = _live(source_section_identifier, "section")
    destination = _live(destination_section_identifier, "section")
    if source["id"] == destination["id"]:
        raise ValueError("The source and destination are the same section.")
    count = len(_pages(onenote.tree(source["id"], "pages")))
    before = len(_pages(onenote.tree(destination["id"], "pages")))
    method = "onenote_merge"
    try:
        onenote.call("merge_sections", source_section_id=source["id"], destination_section_id=destination["id"])
    except RuntimeError as exc:
        # OneNote's merge fails (0x8004201F) for a few minutes after the source section changed; moving the pages
        # through the hierarchy does the same thing and works at once.
        if "8004201F" not in str(exc):
            raise
        method = "page_move"
        target = onenote.tree(destination["id"], "pages")
        pages = _pages(onenote.tree(source["id"], "pages"))
        children = list(target)
        last = _pages(target)[-1] if _pages(target) else None
        insert_at = children.index(last) + 1 if last is not None else len(children)
        for offset, node in enumerate(pages):
            target.insert(insert_at + offset, node)
        _submit_pages(target, {}, set())
    after = before
    for _ in range(20):
        after = len(_pages(onenote.tree(destination["id"], "pages")))
        if after >= before + count:
            break
        time.sleep(0.25)
    if after < before + count:
        raise RuntimeError(f"OneNote merged only {after - before} of {count} pages; the source section was kept.")
    still_there = onenote.find(id=source["id"], item_type="section", fresh=True)
    if still_there is not None and not in_recycle_bin(still_there):
        onenote.call("delete_hierarchy", object_id=source["id"], permanently=False)
    return ok(
        source_section_id=source["id"],
        destination_section_id=destination["id"],
        merged_pages=count,
        method=method,
        source_deleted="moved to the recycle bin",
    )


def _page_ids_under(item: Item) -> list[str]:
    if item["type"] == "page":
        return [item["id"]]
    return [page["id"] for page in of_type(parse_hierarchy(onenote.hierarchy_xml(item["id"], "pages")), "page")]


@tool(DESTRUCTIVE)
def delete_item(identifier: str, permanently: bool = False) -> dict[str, Any]:
    """Delete a page, a section or a section group. permanently=false moves it to OneNote's recycle bin.

    Pages are backed up before a delete; a permanently deleted section has each of its pages backed up.
    """

    item = onenote.resolve(identifier)
    if item["type"] == "notebook":
        raise ValueError("Notebooks are folders on disk; close them with close_notebook and delete the folder yourself.")
    backed_up = []
    if item["type"] == "page" or permanently:
        backed_up = [backup for page_id in _page_ids_under(item) if (backup := snapshot(page_id, "delete_item")) is not None]

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
