"""OneNote XML builders, page readers, and hierarchy helpers."""

from __future__ import annotations

import html
import xml.etree.ElementTree as ET
from dataclasses import dataclass
from typing import Any, Iterable, Literal

from .constants import ONE_NS
from .markup import StyleSet, cdata, html_to_inline, html_to_oe_xml, markdown_to_html, one_t, plain_to_oe_xml
from .reader import ReadFormat, page_title, render_page

ET.register_namespace("one", ONE_NS)


DELETABLE_PAGE_OBJECT_TYPES = {"Outline", "Image", "InkDrawing", "FileAttachment", "InsertedFile", "MediaFile"}
CONTENT_FORMATS = ("plain", "html", "markdown")


def _content_format(content_format: str) -> str:
    key = (content_format or "plain").strip().casefold()
    if key == "md":
        key = "markdown"
    if key not in CONTENT_FORMATS:
        raise ValueError("content_format must be 'plain', 'html', or 'markdown'.")
    return key


def normalize_content(content: str, content_format: str = "plain") -> str:
    """Return a single OneNote inline HTML fragment for plain text, HTML, or Markdown."""

    fmt = _content_format(content_format)
    if fmt == "plain":
        escaped = html.escape(content, quote=False)
        return escaped.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "<br/>")
    if fmt == "html":
        return html_to_inline(content)
    return html_to_inline(markdown_to_html(content))


def attr(value: str) -> str:
    return html.escape(value, quote=True)


def content_to_oe_xml(
    content: str,
    content_format: Literal["plain", "html", "markdown", "md"] = "plain",
    styles: StyleSet | None = None,
) -> str:
    fmt = _content_format(content_format)
    if fmt == "plain":
        return plain_to_oe_xml(content)
    if fmt == "html":
        return html_to_oe_xml(content, styles)
    return html_to_oe_xml(markdown_to_html(content), styles)


def build_outline_xml(
    content: str,
    *,
    content_format: Literal["plain", "html", "markdown", "md"] = "plain",
    object_id: str | None = None,
    x: float | None = None,
    y: float | None = None,
    styles: StyleSet | None = None,
) -> str:
    object_attr = f' objectID="{attr(object_id)}"' if object_id else ""
    position = ""
    if x is not None or y is not None:
        px = 36.0 if x is None else float(x)
        py = 86.0 if y is None else float(y)
        position = f'<one:Position x="{px:.2f}" y="{py:.2f}" z="0"/>'
    return (
        f"<one:Outline{object_attr}>"
        f"{position}"
        f"<one:OEChildren>{content_to_oe_xml(content, content_format, styles)}</one:OEChildren>"
        "</one:Outline>"
    )


def build_title_xml(title: str) -> str:
    return f"<one:Title><one:OE>{one_t(html.escape(title, quote=False))}</one:OE></one:Title>"


def build_page_update_xml(
    page_id: str,
    *,
    title: str | None = None,
    content: str | None = None,
    content_format: str = "plain",
    x: float | None = None,
    y: float | None = None,
    native_styles: bool = True,
) -> str:
    """Page update XML. native_styles uses OneNote heading/quote/code styles and To Do tags."""

    styles = StyleSet() if native_styles else None
    outline = ""
    if content is not None and content != "":
        outline = build_outline_xml(content, content_format=content_format, x=x, y=y, styles=styles)
    parts = [f'<one:Page xmlns:one="{ONE_NS}" ID="{attr(page_id)}">']
    if styles is not None:
        parts.append(styles.defs_xml())
    if title is not None:
        parts.append(build_title_xml(title))
    parts.append(outline)
    parts.append("</one:Page>")
    return "".join(parts)


def build_image_page_update_xml(
    page_id: str,
    *,
    image_base64: str,
    image_format: str,
    x: float | None = None,
    y: float | None = None,
    width: float | None = None,
    height: float | None = None,
) -> str:
    size = ""
    if width is not None and height is not None:
        size = f'<one:Size width="{float(width):.2f}" height="{float(height):.2f}"/>'
    position = ""
    if x is not None or y is not None:
        px = 36.0 if x is None else float(x)
        py = 86.0 if y is None else float(y)
        position = f'<one:Position x="{px:.2f}" y="{py:.2f}" z="0"/>'
    return (
        f'<one:Page xmlns:one="{ONE_NS}" ID="{attr(page_id)}">'
        "<one:Outline>"
        f"{position}"
        "<one:OEChildren><one:OE>"
        f'<one:Image format="{attr(image_format.lower())}">'
        f"{size}<one:Data>{image_base64}</one:Data>"
        "</one:Image>"
        "</one:OE></one:OEChildren>"
        "</one:Outline>"
        "</one:Page>"
    )


PLACED_PAGE_OBJECTS = {"Outline", "Image", "InkDrawing", "InsertedFile", "MediaFile"}


def content_bottom(xml: str) -> float | None:
    """Lowest edge (in points) of the content blocks on a page, or None for an empty page."""

    bottom = None
    for node in parse_xml(xml):
        if local_name(node.tag) not in PLACED_PAGE_OBJECTS:
            continue
        position = node.find(f"{{{ONE_NS}}}Position")
        size = node.find(f"{{{ONE_NS}}}Size")
        if position is None or size is None:
            continue
        try:
            edge = float(position.attrib.get("y", "0")) + float(size.attrib.get("height", "0"))
        except ValueError:
            continue
        bottom = edge if bottom is None else max(bottom, edge)
    return bottom


def local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1] if "}" in tag else tag


def parse_xml(xml: str) -> ET.Element:
    return ET.fromstring(xml.encode("utf-8"))


def text_from_page_xml(xml: str, fmt: ReadFormat = "text", *, include_ocr: bool = True) -> str:
    """Readable page body: list markers, tables, and image OCR text included."""

    return render_page(xml, fmt, include_ocr=include_ocr, include_title=False)


def title_from_page_xml(xml: str) -> str | None:
    return page_title(xml)


def collect_page_objects(xml: str) -> list[dict[str, Any]]:
    root = parse_xml(xml)
    objects = []
    content_without_own_id = {"Image", "FileAttachment", "InsertedFile", "MediaFile"}

    def walk(
        node: ET.Element,
        container_object_id: str | None = None,
        deletable_container_id: str | None = None,
        in_title: bool = False,
    ) -> None:
        kind = local_name(node.tag)
        next_in_title = in_title or kind == "Title"
        object_id = node.attrib.get("objectID") or node.attrib.get("ID")
        next_container_id = object_id or container_object_id
        delete_supported = kind in DELETABLE_PAGE_OBJECT_TYPES and bool(object_id)
        next_deletable_container_id = object_id if delete_supported else deletable_container_id

        if not next_in_title and kind != "Page" and (object_id or kind in content_without_own_id):
            record: dict[str, Any] = {"type": kind}
            if object_id:
                record["object_id"] = object_id
            elif container_object_id:
                record["container_object_id"] = container_object_id
            if container_object_id and object_id != container_object_id:
                record["parent_object_id"] = container_object_id
            record["delete_supported"] = delete_supported
            if delete_supported and object_id:
                record["delete_object_id"] = object_id
            elif deletable_container_id:
                record["delete_object_id"] = deletable_container_id
            callback = node.find(f"{{{ONE_NS}}}CallbackID")
            if callback is not None and callback.attrib.get("callbackID"):
                record["callback_id"] = callback.attrib["callbackID"]
            elif "callbackID" in node.attrib:
                record["callback_id"] = node.attrib["callbackID"]
            for name in ("preferredName", "alt"):
                if node.attrib.get(name):
                    record[name] = node.attrib[name]
            if "format" in node.attrib:
                record["format"] = node.attrib["format"]
            objects.append(record)

        for child in list(node):
            walk(child, next_container_id, next_deletable_container_id, next_in_title)

    walk(root)
    return objects


@dataclass
class HierarchyItem:
    type: str
    id: str
    name: str
    path: str
    level: int
    parent_id: str | None
    parent_name: str | None
    notebook_name: str | None
    section_name: str | None
    attributes: dict[str, str]

    def as_dict(self) -> dict[str, Any]:
        data = {
            "type": self.type,
            "id": self.id,
            "name": self.name,
            "path": self.path,
            "level": self.level,
            "parent_id": self.parent_id,
            "parent_name": self.parent_name,
            "notebook_name": self.notebook_name,
            "section_name": self.section_name,
        }
        data.update(self.attributes)
        return data


TYPE_MAP = {
    "Notebook": "notebook",
    "SectionGroup": "section_group",
    "Section": "section",
    "Page": "page",
}


def parse_hierarchy(xml: str) -> list[dict[str, Any]]:
    root = parse_xml(xml)
    items: list[HierarchyItem] = []

    def walk(
        node: ET.Element,
        ancestors: list[str],
        *,
        parent_id: str | None,
        parent_name: str | None,
        notebook_name: str | None,
        section_name: str | None,
        level: int,
    ) -> None:
        node_type = local_name(node.tag)
        if node_type in TYPE_MAP:
            name = node.attrib.get("name") or node.attrib.get("nickname") or "(untitled)"
            object_id = node.attrib.get("ID", "")
            path_parts = ancestors + [name]
            current_notebook = notebook_name
            current_section = section_name
            if node_type == "Notebook":
                current_notebook = name
            elif node_type == "Section":
                current_section = name
            attributes = {}
            for key, value in node.attrib.items():
                if key in {"ID", "name"}:
                    continue
                if key == "path":
                    attributes["onenote_path"] = value
                else:
                    attributes[key] = value
            items.append(
                HierarchyItem(
                    type=TYPE_MAP[node_type],
                    id=object_id,
                    name=name,
                    path="/".join(path_parts),
                    level=level,
                    parent_id=parent_id,
                    parent_name=parent_name,
                    notebook_name=current_notebook,
                    section_name=current_section,
                    attributes=attributes,
                )
            )
            next_parent_id = object_id
            next_parent_name = name
            next_ancestors = path_parts
            next_notebook = current_notebook
            next_section = current_section
            next_level = level + 1
        else:
            next_parent_id = parent_id
            next_parent_name = parent_name
            next_ancestors = ancestors
            next_notebook = notebook_name
            next_section = section_name
            next_level = level

        for child in list(node):
            walk(
                child,
                next_ancestors,
                parent_id=next_parent_id,
                parent_name=next_parent_name,
                notebook_name=next_notebook,
                section_name=next_section,
                level=next_level,
            )

    walk(
        root,
        [],
        parent_id=None,
        parent_name=None,
        notebook_name=None,
        section_name=None,
        level=0,
    )
    return [item.as_dict() for item in items]


def filter_items(items: Iterable[dict[str, Any]], item_type: str) -> list[dict[str, Any]]:
    return [item for item in items if item.get("type") == item_type]


def resolve_item(items: Iterable[dict[str, Any]], identifier: str, item_type: str | None = None) -> dict[str, Any]:
    candidates = [item for item in items if item_type is None or item.get("type") == item_type]
    type_label = item_type or "object"
    for item in candidates:
        if item.get("id") == identifier:
            return item
    lowered = identifier.casefold()
    path_exact = [item for item in candidates if item.get("path", "").casefold() == lowered]
    if len(path_exact) == 1:
        return path_exact[0]
    if len(path_exact) > 1:
        paths = ", ".join(item["path"] for item in path_exact[:10])
        raise ValueError(f"Ambiguous {type_label} identifier '{identifier}'. Use an ID or exact path. Matches: {paths}")

    name_exact = [item for item in candidates if item.get("name", "").casefold() == lowered]
    if len(name_exact) == 1:
        return name_exact[0]
    if len(name_exact) > 1:
        paths = ", ".join(item["path"] for item in name_exact[:10])
        raise ValueError(f"Ambiguous {type_label} identifier '{identifier}'. Use an ID or exact path. Matches: {paths}")
    raise ValueError(
        f"No {type_label} found for '{identifier}'. Use an ID or exact path from list_hierarchy, "
        "list_sections, or list_pages."
    )
