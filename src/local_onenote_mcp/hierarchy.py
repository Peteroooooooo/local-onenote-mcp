"""OneNote hierarchy XML: flatten notebooks, section groups, sections and pages into plain dicts."""

from __future__ import annotations

import re
import xml.etree.ElementTree as ET
from typing import Any, Iterable


TYPE_MAP = {"Notebook": "notebook", "SectionGroup": "section_group", "Section": "section", "Page": "page"}
TAG_FOR_TYPE = {value: key for key, value in TYPE_MAP.items()}
ONENOTE_ID = re.compile(r"^\{[0-9A-Fa-f-]{36}\}\{\d+\}\{[0-9A-Za-z]+\}$")
RECYCLE_BIN = "OneNote_RecycleBin"

Item = dict[str, Any]


def local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def parse_xml(xml: str) -> ET.Element:
    return ET.fromstring(xml.encode("utf-8"))


def parse_hierarchy(xml: str | ET.Element) -> list[Item]:
    """Every notebook/section group/section/page as a dict with its slash-separated path."""

    root = parse_xml(xml) if isinstance(xml, str) else xml
    items: list[Item] = []

    def walk(node: ET.Element, parent: Item | None) -> None:
        kind = TYPE_MAP.get(local_name(node.tag))
        current = parent
        if kind:
            name = node.attrib.get("name") or node.attrib.get("nickname") or "(untitled)"
            path = f"{parent['path']}/{name}" if parent else name
            item: Item = {
                "type": kind,
                "id": node.attrib.get("ID", ""),
                "name": name,
                "path": path,
                "level": parent["level"] + 1 if parent else 0,
                "parent_id": parent["id"] if parent else None,
                "notebook_name": name if kind == "notebook" else (parent or {}).get("notebook_name"),
                "section_name": name if kind == "section" else (parent or {}).get("section_name"),
            }
            for key, value in node.attrib.items():
                if key not in {"ID", "name"}:
                    item["onenote_path" if key == "path" else key] = value
            items.append(item)
            current = item
        for child in node:
            walk(child, current)

    walk(root, None)
    return items


def of_type(items: Iterable[Item], item_type: str) -> list[Item]:
    return [item for item in items if item.get("type") == item_type]


def in_recycle_bin(item: Item) -> bool:
    if item.get("isInRecycleBin") == "true" or item.get("isRecycleBin") == "true":
        return True
    return RECYCLE_BIN in item.get("path", "").split("/")


def without_recycle_bin(items: Iterable[Item]) -> list[Item]:
    return [item for item in items if not in_recycle_bin(item)]


def normal_path(value: str) -> str:
    """Case-folded path with spaces around each '/' and at the ends ignored, so 'Nb / Sec/' finds 'Nb/Sec'."""

    return "/".join(part.strip() for part in value.strip().strip("/").split("/")).casefold()


def resolve_item(items: Iterable[Item], identifier: str, item_type: str | None = None) -> Item:
    """Match by ID, then exact path, then unique name (both case-insensitive)."""

    candidates = [item for item in items if item_type is None or item.get("type") == item_type]
    label = item_type or "object"
    for item in candidates:
        if item.get("id") == identifier:
            return item
    wanted = normal_path(identifier)
    for key in ("path", "name"):
        matches = [item for item in candidates if (normal_path(item.get(key, "")) if key == "path" else item.get(key, "").strip().casefold()) == wanted]
        if len(matches) == 1:
            return matches[0]
        if len(matches) > 1:
            live = without_recycle_bin(matches)
            if len(live) == 1:
                return live[0]
            listed = "; ".join(f"{item['path']} (ID {item['id']})" for item in matches[:10])
            raise ValueError(f"Ambiguous {label} identifier '{identifier}': {len(matches)} matches. Use one of these IDs: {listed}")
    raise ValueError(f"No {label} found for '{identifier}'. Use an ID or exact path from list_hierarchy.")


def safe_leaf_name(name: str) -> str:
    cleaned = re.sub(r"\s+", " ", re.sub(r'[<>:"/\\|?*\x00-\x1f]', " ", name)).strip()
    if not cleaned:
        raise ValueError("Name cannot be empty.")
    return cleaned


def child_path(parent_path: str, child_name: str) -> str:
    name = child_name.replace("\\", "/").strip("/")
    if name.lower().endswith(".one"):
        name = name[:-4]
    return f"{parent_path}/{name}" if name else parent_path
