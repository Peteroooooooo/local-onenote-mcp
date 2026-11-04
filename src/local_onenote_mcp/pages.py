"""Page-level helpers shared by several tools: snapshots, copies/restores, and placement."""

from __future__ import annotations

import base64
import copy
import xml.etree.ElementTree as ET
from pathlib import Path

from .app import backups, onenote
from .backup import BLOB_ATTR, FILE_TAGS
from .config import data_dir
from .constants import ONE
from .hierarchy import parse_xml, safe_leaf_name
from .pagexml import body_object_ids, content_bottom, definitions
from .reader import page_title

BLOCK_GAP = 18.0
BINARY_TAGS = {ONE + tag for tag in ("Image", "InkDrawing", "InkParagraph", "InkWord", "InsertedFile", "MediaFile")}


def below_existing(page_id: str, x: float | None, y: float | None) -> tuple[float | None, float | None]:
    """Keep caller coordinates; otherwise place a new block under the page's current content."""

    if x is not None or y is not None:
        return x, y
    bottom = content_bottom(onenote.page_xml(page_id))
    return (None, None) if bottom is None else (36.0, bottom + BLOCK_GAP)


def file_bytes(page_id: str, node: ET.Element) -> bytes | None:
    """Bytes of an image or attached file node; attachments come from OneNote's temp copy (pathCache)."""

    callback = node.find(ONE + "CallbackID")
    if callback is not None and callback.attrib.get("callbackID"):
        try:
            data = onenote.call("get_binary_page_content", page_id=page_id, callback_id=callback.attrib["callbackID"])
            return base64.b64decode(data["base64"])
        except RuntimeError:
            pass
    cached = node.attrib.get("pathCache")
    if cached and Path(cached).is_file():
        return Path(cached).read_bytes()
    return None


def snapshot(page_id: str, reason: str) -> int | None:
    """Save the page (with images, ink and attachments) before a destructive change."""

    if not backups.enabled:
        return None
    xml = onenote.page_xml(page_id, "all")
    item = onenote.resolve(page_id, "page")
    return backups.save(
        xml,
        page_id=page_id,
        section_id=item.get("parent_id") or "",
        path=item.get("path", ""),
        title=page_title(xml) or item.get("name", ""),
        reason=reason,
        read_file=lambda node: file_bytes(page_id, node),
    )


def snapshot_blocks(page_id: str, page: ET.Element, outline_ids: set[str], reason: str) -> int | None:
    """Save only some outlines of a page (its basic XML) before they are edited.

    Falls back to a full snapshot when those outlines hold images, ink or files,
    whose bytes the basic XML does not carry.
    """

    if not backups.enabled:
        return None
    outlines = [node for node in page.findall(ONE + "Outline") if node.attrib.get("objectID") in outline_ids]
    if not outlines:
        return None
    if any(node.tag in BINARY_TAGS for outline in outlines for node in outline.iter()):
        return snapshot(page_id, reason)
    saved = ET.Element(page.tag, {"ID": page_id})
    saved.extend(copy.deepcopy(node) for node in [*definitions(page), *outlines])
    xml = ET.tostring(saved, encoding="unicode")
    item = onenote.resolve(page_id, "page")
    title = page_title(ET.tostring(page, encoding="unicode")) or item.get("name", "")
    return backups.save(
        xml, page_id=page_id, section_id=item.get("parent_id") or "", path=item.get("path", ""), title=title, reason=reason, kind="blocks"
    )


def restore_blocks(page_id: str, saved: ET.Element) -> None:
    """Put the outlines of a "blocks" snapshot back into their page.

    Outlines and paragraphs that no longer exist are written as new objects; an
    outline that still exists is replaced as a whole, so paragraphs added since
    the snapshot go away again.
    """

    current = parse_xml(onenote.page_xml(page_id))
    live = {node.attrib.get("objectID") for node in current.iter()}
    outlines = [copy.deepcopy(node) for node in saved.findall(ONE + "Outline")]
    for outline in outlines:
        for node in outline.iter():
            if node.attrib.get("objectID") not in live:
                node.attrib.pop("objectID", None)
    defs = _merged_definitions(current, saved, outlines)
    restored = ET.Element(ONE + "Page", {"ID": page_id})
    restored.extend(defs + outlines)
    onenote.update_page(restored)


def _def_key(node: ET.Element) -> tuple[str, ...]:
    return (node.tag, node.attrib.get("name", ""), node.attrib.get("type", ""), node.attrib.get("symbol", ""))


def _merged_definitions(current: ET.Element, saved: ET.Element, outlines: list[ET.Element]) -> list[ET.Element]:
    """The current page's style and tag definitions, plus any the saved outlines need, with their indexes remapped.

    OneNote renumbers definitions when a page changes, so a saved index may now mean another style.
    """

    defs = [copy.deepcopy(node) for node in definitions(current)]
    remap: dict[str, dict[str, str]] = {ONE + "TagDef": {}, ONE + "QuickStyleDef": {}}
    for node in definitions(saved):
        same_kind = [d for d in defs if d.tag == node.tag]
        match = next((d for d in same_kind if _def_key(d) == _def_key(node)), None)
        if match is None:
            match = copy.deepcopy(node)
            match.set("index", str(1 + max((int(d.attrib.get("index", -1)) for d in same_kind), default=-1)))
            defs.insert(max((defs.index(d) + 1 for d in same_kind), default=0 if node.tag == ONE + "TagDef" else len(defs)), match)
        remap[node.tag][node.attrib.get("index", "")] = match.attrib["index"]
    for outline in outlines:
        for node in outline.iter():
            if "quickStyleIndex" in node.attrib:
                node.set("quickStyleIndex", remap[ONE + "QuickStyleDef"].get(node.attrib["quickStyleIndex"], node.attrib["quickStyleIndex"]))
            if node.tag == ONE + "Tag":
                node.set("index", remap[ONE + "TagDef"].get(node.attrib.get("index", ""), node.attrib.get("index", "")))
    return defs


def _materialize(digest: str, name: str, data: bytes) -> Path:
    folder = data_dir() / "attachments" / digest[:16]
    folder.mkdir(parents=True, exist_ok=True)
    target = folder / safe_leaf_name(name or "attachment")
    if not target.exists():
        target.write_bytes(data)
    return target


def portable(page: ET.Element, page_id: str, files: dict[str, bytes] | None = None) -> ET.Element:
    """A copy of a full page tree that can be written into another (new) page.

    Object IDs and callback references are removed so OneNote creates fresh
    objects; attached files point at a local copy of their bytes.
    """

    root = copy.deepcopy(page)
    parents = {child: parent for parent in root.iter() for child in parent}
    for node in list(root.iter()):
        node.attrib.pop("objectID", None)
        if node.tag == ONE + "CallbackID":
            parents[node].remove(node)
        elif node.tag in FILE_TAGS:
            digest = node.attrib.pop(BLOB_ATTR, None)
            cached = node.attrib.pop("pathCache", None)
            name = node.attrib.get("preferredName", "attachment")
            if digest and files and digest in files:
                node.set("pathSource", str(_materialize(digest, name, files[digest])))
            elif cached and Path(cached).is_file():
                node.set("pathSource", cached)
            elif not (node.attrib.get("pathSource") and Path(node.attrib["pathSource"]).is_file()):
                placeholder = ET.Element(ONE + "T")
                placeholder.text = f"[attachment not available: {name}]"
                parent = parents[node]
                parent.insert(list(parent).index(node), placeholder)
                parent.remove(node)
    for key in list(root.attrib):
        if key not in {"ID", "lang"}:
            del root.attrib[key]
    root.set("ID", page_id)
    return root


def write_copy(page: ET.Element, section_id: str, files: dict[str, bytes] | None = None) -> str:
    """Create a new page in section_id holding the content of a full page tree. Returns the new page ID."""

    new_id = onenote.new_page(section_id, style=2)
    try:
        onenote.update_page(portable(page, new_id, files))
    except Exception:
        onenote.call("delete_hierarchy", object_id=new_id, permanently=True)
        raise
    return new_id


def replace_contents(page_id: str, source: ET.Element, files: dict[str, bytes] | None = None) -> list[str]:
    """Make an existing page show source's body and title: write the new content first, then drop the old objects."""

    old_ids = body_object_ids(parse_xml(onenote.page_xml(page_id)))
    onenote.update_page(portable(source, page_id, files), force=True)
    for object_id in old_ids:
        onenote.delete_object(page_id, object_id)
    return old_ids
