"""Build OneNote page XML for updates, and inspect page XML returned by OneNote."""

from __future__ import annotations

import html
import xml.etree.ElementTree as ET
from typing import Any

from .constants import ONE, ONE_NS
from .hierarchy import local_name, parse_xml
from .markup import StyleSet, html_to_oe_xml, markdown_to_html, one_t, plain_to_oe_xml

CONTENT_FORMATS = ("plain", "html", "markdown")
BODY_TYPES = {"Outline", "Image", "InkDrawing", "InsertedFile", "MediaFile"}
DEFINITION_TAGS = {ONE + "TagDef", ONE + "QuickStyleDef"}


def content_format(value: str) -> str:
    key = (value or "plain").strip().casefold()
    key = "markdown" if key == "md" else key
    if key not in CONTENT_FORMATS:
        raise ValueError("content_format must be 'plain', 'html', or 'markdown'.")
    return key


def content_to_oe_xml(content: str, fmt: str = "plain", styles: StyleSet | None = None) -> str:
    fmt = content_format(fmt)
    if fmt == "plain":
        return plain_to_oe_xml(content)
    return html_to_oe_xml(content if fmt == "html" else markdown_to_html(content), styles)


def attr(value: str) -> str:
    return html.escape(value, quote=True)


def _position(x: float | None, y: float | None) -> str:
    if x is None and y is None:
        return ""
    return f'<one:Position x="{36.0 if x is None else float(x):.2f}" y="{86.0 if y is None else float(y):.2f}" z="0"/>'


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
    """A page update that sets the title and/or adds one new outline."""

    styles = StyleSet() if native_styles else None
    outline = ""
    body = content_to_oe_xml(content, content_format, styles) if content and content.strip() else ""
    if body:
        outline = f"<one:Outline>{_position(x, y)}<one:OEChildren>{body}</one:OEChildren></one:Outline>"
    if title is not None:
        # A page title is one line; a newline in it would leave a broken second title paragraph.
        title = " ".join(title.split())
    title_xml = f"<one:Title><one:OE>{one_t(html.escape(title, quote=False))}</one:OE></one:Title>" if title is not None else ""
    defs = styles.defs_xml() if styles else ""
    return f'<one:Page xmlns:one="{ONE_NS}" ID="{attr(page_id)}">{defs}{title_xml}{outline}</one:Page>'


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
    size = f'<one:Size width="{float(width):.2f}" height="{float(height):.2f}"/>' if width is not None and height is not None else ""
    return (
        f'<one:Page xmlns:one="{ONE_NS}" ID="{attr(page_id)}"><one:Outline>{_position(x, y)}<one:OEChildren><one:OE>'
        f'<one:Image format="{attr(image_format.lower())}">{size}<one:Data>{image_base64}</one:Data></one:Image>'
        "</one:OE></one:OEChildren></one:Outline></one:Page>"
    )


def build_file_page_update_xml(page_id: str, *, path: str, name: str, x: float | None = None, y: float | None = None) -> str:
    return (
        f'<one:Page xmlns:one="{ONE_NS}" ID="{attr(page_id)}"><one:Outline>{_position(x, y)}<one:OEChildren><one:OE>'
        f'<one:InsertedFile pathSource="{attr(path)}" preferredName="{attr(name)}"/>'
        "</one:OE></one:OEChildren></one:Outline></one:Page>"
    )


def content_bottom(xml: str | ET.Element) -> float | None:
    """Lowest edge (in points) of the content blocks on a page, or None for an empty page."""

    root = parse_xml(xml) if isinstance(xml, str) else xml
    bottom = None
    for node in root:
        if local_name(node.tag) not in BODY_TYPES:
            continue
        position, size = node.find(ONE + "Position"), node.find(ONE + "Size")
        if position is None or size is None:
            continue
        try:
            edge = float(position.attrib.get("y", "0")) + float(size.attrib.get("height", "0"))
        except ValueError:
            continue
        bottom = edge if bottom is None else max(bottom, edge)
    return bottom


def collect_page_objects(xml: str | ET.Element) -> list[dict[str, Any]]:
    """Blocks, images, ink and attachments outside the title, with the IDs other tools need.

    Paragraphs, tables and cells are left out (get_page(include_paragraphs=true) lists those).
    delete_object_id is the nearest object DeletePageContent accepts (outlines and
    top-level images/ink/files); callback_id downloads binary content. Top-level
    blocks carry x/y/width/height in points.
    """

    root = parse_xml(xml) if isinstance(xml, str) else xml
    objects: list[dict[str, Any]] = []

    def walk(node: ET.Element, parent_id: str | None, deletable_id: str | None, top: bool) -> None:
        kind = local_name(node.tag)
        if kind == "Title":
            return
        object_id = node.attrib.get("objectID")
        deletable = top and kind in BODY_TYPES and bool(object_id)
        if kind in BODY_TYPES:
            record: dict[str, Any] = {"type": kind}
            if object_id:
                record["object_id"] = object_id
            if parent_id:
                record["parent_object_id"] = parent_id
            record["delete_supported"] = deletable
            if deletable or deletable_id:
                record["delete_object_id"] = object_id if deletable else deletable_id
            callback = node.find(ONE + "CallbackID")
            if callback is not None and callback.attrib.get("callbackID"):
                record["callback_id"] = callback.attrib["callbackID"]
            for name in ("format", "preferredName", "alt"):
                if node.attrib.get(name):
                    record[name] = node.attrib[name]
            position, size = node.find(ONE + "Position"), node.find(ONE + "Size")
            if top and position is not None and size is not None:
                for key, source in (("x", position), ("y", position), ("width", size), ("height", size)):
                    try:
                        record[key] = round(float(source.attrib[key]), 1)
                    except (KeyError, ValueError):
                        pass
            objects.append(record)
            parent_id = object_id or parent_id
        for child in node:
            walk(child, parent_id, object_id if deletable else deletable_id, False)

    for child in root:
        walk(child, None, None, True)
    return objects


def definitions(root: ET.Element) -> list[ET.Element]:
    return [node for node in root if node.tag in DEFINITION_TAGS]


def body_object_ids(root: ET.Element) -> list[str]:
    return [node.attrib["objectID"] for node in root if local_name(node.tag) in BODY_TYPES and node.attrib.get("objectID")]
