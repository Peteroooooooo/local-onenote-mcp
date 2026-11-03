"""Paragraph-level reading and editing of an existing OneNote page.

A paragraph is one one:OE in the page body. Edits change the parsed page tree in
place; build_update() then returns XML holding only the outlines that changed,
with the page's style and tag definitions, so everything else on the page
(images, ink, attachments, untouched outlines) is never resubmitted.
"""

from __future__ import annotations

import copy
import html
import re
import xml.etree.ElementTree as ET
from dataclasses import dataclass
from typing import Any, Literal

from pydantic import BaseModel, Field

from .constants import ONE, ONE_NS
from .markup import StyleSet, html_to_inline, markdown_to_html
from .pagexml import attr, content_format, content_to_oe_xml, definitions
from .reader import inline_text

TODO_NAMES = {"to do", "待办事项", "待办"}


class EditOperation(BaseModel):
    """One change to a page. Paragraph IDs come from get_page(include_paragraphs=true)."""

    op: Literal["replace_text", "set_text", "insert_after", "insert_before", "delete", "set_todo", "add_row", "delete_row"] = Field(
        description=(
            "replace_text: find/replace across the page body, formatting kept. "
            "set_text: replace one paragraph's text, keeping its bullet, number, to-do box and style. "
            "insert_after / insert_before: add content (any length, any format) next to a paragraph. "
            "delete: remove a paragraph and its indented children. "
            "set_todo: tick or untick a paragraph's to-do box, adding one if it has none. "
            "add_row: add a table row after the row holding paragraph_id (or at the end, given the table's own paragraph). "
            "delete_row: remove the table row holding paragraph_id."
        )
    )
    paragraph_id: str = Field("", description="Target paragraph for every op except replace_text. Table cells have their own paragraphs.")
    find: str = Field("", description="replace_text: exact text to find (case-sensitive).")
    replace: str = Field("", description="replace_text: replacement text.")
    max_count: int = Field(0, description="replace_text: replace at most this many matches; 0 = all.")
    content: str = Field("", description="set_text / insert_*: the new content.")
    cells: list[str] = Field(default_factory=list, description="add_row: text of each cell, left to right; missing cells stay empty.")
    content_format: Literal["plain", "html", "markdown"] = Field("markdown", description="Format of content and cells.")
    done: bool = Field(True, description="set_todo: true = ticked.")


@dataclass
class _Located:
    oe: ET.Element
    parent: ET.Element
    outline: ET.Element
    table: ET.Element | None = None
    row: ET.Element | None = None


def is_todo_def(node: ET.Element) -> bool:
    return node.attrib.get("type") == "0" or node.attrib.get("name", "").casefold() in TODO_NAMES


class PageEditor:
    def __init__(self, page: ET.Element) -> None:
        self.page = page
        self.styles = StyleSet(definitions(page))
        self.tag_names = {node.attrib.get("index"): node for node in page if node.tag == ONE + "TagDef"}
        self.quick_names = {node.attrib.get("index"): node.attrib.get("name", "") for node in page if node.tag == ONE + "QuickStyleDef"}
        self.touched: list[ET.Element] = []
        self._index: dict[str, _Located] = {}
        for outline in page.findall(ONE + "Outline"):
            self._walk(outline.find(ONE + "OEChildren"), outline)

    def _walk(self, container: ET.Element | None, outline: ET.Element, table: ET.Element | None = None, row: ET.Element | None = None) -> None:
        if container is None:
            return
        for oe in container.findall(ONE + "OE"):
            if oe.attrib.get("objectID"):
                self._index[oe.attrib["objectID"]] = _Located(oe, container, outline, table, row)
            self._walk(oe.find(ONE + "OEChildren"), outline, table, row)
            inner = oe.find(ONE + "Table")
            if inner is not None:
                for inner_row in inner.findall(ONE + "Row"):
                    for cell in inner_row.findall(ONE + "Cell"):
                        self._walk(cell.find(ONE + "OEChildren"), outline, inner, inner_row)

    # -- reading ---------------------------------------------------------

    def paragraphs(self) -> list[dict[str, Any]]:
        result: list[dict[str, Any]] = []
        for outline in self.page.findall(ONE + "Outline"):
            self._describe(outline.find(ONE + "OEChildren"), 0, result, {})
        return result

    def _describe(self, container: ET.Element | None, depth: int, out: list[dict[str, Any]], cell: dict[str, Any]) -> None:
        if container is None:
            return
        for oe in container.findall(ONE + "OE"):
            record: dict[str, Any] = {"id": oe.attrib.get("objectID", ""), "depth": depth, **cell}
            text_node = oe.find(ONE + "T")
            table = oe.find(ONE + "Table")
            if text_node is not None:
                record["text"] = inline_text(text_node.text).replace("\xa0", " ")
            elif table is not None:
                rows = table.findall(ONE + "Row")
                columns = max((len(row.findall(ONE + "Cell")) for row in rows), default=0)
                record["kind"] = f"table ({len(rows)} rows x {columns} columns)"
            else:
                kinds = [child.tag.removeprefix(ONE) for child in oe if child.tag.removeprefix(ONE) in {"Image", "InkDrawing", "InkWord", "InsertedFile", "MediaFile"}]
                record["kind"] = kinds[0].lower() if kinds else "empty"
            style = self.quick_names.get(oe.attrib.get("quickStyleIndex"))
            if style and style not in {"p", "PageTitle"}:
                record["style"] = style
            list_node = oe.find(ONE + "List")
            if list_node is not None:
                record["list"] = "number" if list_node.find(ONE + "Number") is not None else "bullet"
            todo = self._todo(oe)
            if todo is not None:
                record["todo"] = todo.attrib.get("completed") == "true"
            out.append(record)
            self._describe(oe.find(ONE + "OEChildren"), depth + 1, out, cell)
            if table is not None:
                for r, row in enumerate(table.findall(ONE + "Row")):
                    for c, cell_node in enumerate(row.findall(ONE + "Cell")):
                        where = {"table": record["id"], "row": r, "column": c}
                        self._describe(cell_node.find(ONE + "OEChildren"), depth + 1, out, where)

    def _todo(self, oe: ET.Element) -> ET.Element | None:
        for tag in oe.findall(ONE + "Tag"):
            definition = self.tag_names.get(tag.attrib.get("index"))
            if definition is not None and is_todo_def(definition):
                return tag
        return None

    # -- editing ---------------------------------------------------------

    def apply(self, op: EditOperation) -> dict[str, Any]:
        if op.op == "replace_text":
            if not op.find:
                raise ValueError("replace_text needs 'find'.")
            count = self.replace_text(op.find, op.replace, op.max_count)
            if count == 0:
                raise ValueError(
                    f"Text '{op.find}' was not found. Matches must lie inside one formatting run; "
                    "use set_text on the paragraph for text that spans bold/italic boundaries."
                )
            return {"op": op.op, "replacements": count}
        located = self._locate(op.paragraph_id)
        if op.op == "set_text":
            self.set_text(located, op.content, op.content_format)
        elif op.op in {"insert_after", "insert_before"}:
            added = self.insert(located, op.content, op.content_format, before=op.op == "insert_before")
            return {"op": op.op, "paragraph_id": op.paragraph_id, "inserted_paragraphs": added}
        elif op.op == "delete":
            self.delete(located)
        elif op.op == "set_todo":
            self.set_todo(located, op.done)
        elif op.op == "add_row":
            self.add_row(located, op.cells, op.content_format)
        elif op.op == "delete_row":
            self.delete_row(located)
        return {"op": op.op, "paragraph_id": op.paragraph_id}

    def _locate(self, paragraph_id: str) -> _Located:
        located = self._index.get(paragraph_id)
        attached = located is not None and located.oe in list(located.parent)
        if attached and located.row is not None:
            attached = located.row in list(located.table)
        if not attached:
            raise ValueError(f"Paragraph '{paragraph_id}' is not on this page (deleted earlier in this edit, or a wrong ID).")
        return located

    def _touch(self, outline: ET.Element) -> None:
        if outline not in self.touched:
            self.touched.append(outline)

    def replace_text(self, find: str, replace: str, max_count: int = 0) -> int:
        total = 0
        for outline in self.page.findall(ONE + "Outline"):
            for node in outline.iter(ONE + "T"):
                limit = max_count - total if max_count else 0
                if max_count and limit <= 0:
                    return total
                new_text, count = _replace_in_fragment(node.text or "", find, replace, limit)
                if count:
                    node.text = new_text
                    total += count
                    self._touch(outline)
        return total

    def set_text(self, located: _Located, content: str, fmt: str) -> None:
        text_node = located.oe.find(ONE + "T")
        if text_node is None:
            raise ValueError("This paragraph holds an image, table or file, not text; use insert/delete instead.")
        text_node.text = _inline(content, fmt)
        self._touch(located.outline)

    def insert(self, located: _Located, content: str, fmt: str, before: bool) -> int:
        body = content_to_oe_xml(content, fmt, self.styles)
        new_nodes = list(ET.fromstring(f'<one:OEChildren xmlns:one="{ONE_NS}">{body}</one:OEChildren>'))
        if not new_nodes:
            raise ValueError("content is empty.")
        position = list(located.parent).index(located.oe) + (0 if before else 1)
        for offset, node in enumerate(new_nodes):
            located.parent.insert(position + offset, node)
        self._touch(located.outline)
        return len(new_nodes)

    def delete(self, located: _Located) -> None:
        if located.table is not None and len(located.parent.findall(ONE + "OE")) == 1:
            # A table cell must keep one paragraph; empty it instead.
            for child in list(located.oe):
                located.oe.remove(child)
            ET.SubElement(located.oe, ONE + "T").text = ""
        else:
            located.parent.remove(located.oe)
        self._touch(located.outline)

    def add_row(self, located: _Located, cells: list[str], fmt: str) -> None:
        table, anchor = located.table, located.row
        if table is None:
            table, anchor = located.oe.find(ONE + "Table"), None
        if table is None:
            raise ValueError("add_row needs a paragraph inside a table, or the table's own paragraph.")
        rows = table.findall(ONE + "Row")
        row = copy.deepcopy(rows[-1])
        row_cells = row.findall(ONE + "Cell")
        if len(cells) > len(row_cells):
            raise ValueError(f"The table has {len(row_cells)} columns but {len(cells)} cells were given.")
        for node in row.iter():
            node.attrib.pop("objectID", None)
        for index, cell in enumerate(row_cells):
            container = cell.find(ONE + "OEChildren")
            first = container.find(ONE + "OE")
            for extra in container.findall(ONE + "OE")[1:]:
                container.remove(extra)
            for child in list(first):
                first.remove(child)
            ET.SubElement(first, ONE + "T").text = _inline(cells[index], fmt) if index < len(cells) else ""
        children = list(table)
        table.insert(children.index(anchor if anchor is not None else rows[-1]) + 1, row)
        self._touch(located.outline)

    def delete_row(self, located: _Located) -> None:
        if located.table is None or located.row is None:
            raise ValueError("delete_row needs a paragraph inside a table cell.")
        if len(located.table.findall(ONE + "Row")) == 1:
            raise ValueError("This is the table's only row; delete the table's own paragraph instead.")
        located.table.remove(located.row)
        self._touch(located.outline)

    def set_todo(self, located: _Located, done: bool) -> None:
        tag = self._todo(located.oe)
        if tag is None:
            tag = ET.Element(ONE + "Tag", {"index": str(self.styles.todo_tag())})
            located.oe.insert(0, tag)
        tag.set("completed", "true" if done else "false")
        for stale in ("completionDate",):
            tag.attrib.pop(stale, None)
        self._touch(located.outline)

    # -- output ----------------------------------------------------------

    def empty_outlines(self) -> list[str]:
        """Changed outlines that no longer hold any paragraph; they must be deleted, not updated."""

        return [
            outline.attrib["objectID"]
            for outline in self.touched
            if outline.attrib.get("objectID") and not outline.find(ONE + "OEChildren").findall(ONE + "OE")
        ]

    def build_update(self, page_id: str) -> str | None:
        outlines = [outline for outline in self.touched if outline.find(ONE + "OEChildren").findall(ONE + "OE")]
        if not outlines:
            return None
        body = "".join(ET.tostring(outline, encoding="unicode") for outline in outlines)
        return f'<one:Page xmlns:one="{ONE_NS}" ID="{attr(page_id)}">{self.styles.defs_xml()}{body}</one:Page>'


def _inline(content: str, fmt: str) -> str:
    fmt = content_format(fmt)
    if fmt == "plain":
        return html.escape(content, quote=False).replace("\r\n", "\n").replace("\n", " ")
    return html_to_inline(content if fmt == "html" else markdown_to_html(content))


_TAG_SPLIT = re.compile(r"(<[^>]*>)")


def _replace_in_fragment(fragment: str, find: str, replace: str, limit: int) -> tuple[str, int]:
    """Replace text inside the text runs of an inline HTML fragment, never inside tags."""

    total = 0
    parts = _TAG_SPLIT.split(fragment)
    for index, part in enumerate(parts):
        if not part or part.startswith("<"):
            continue
        plain = html.unescape(part)
        if find not in plain:
            continue
        allowed = plain.count(find) if not limit else min(plain.count(find), limit - total)
        if allowed <= 0:
            break
        parts[index] = html.escape(plain.replace(find, replace, allowed), quote=False)
        total += allowed
        if limit and total >= limit:
            break
    return "".join(parts), total
