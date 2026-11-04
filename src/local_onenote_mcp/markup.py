"""Convert plain text, HTML, and Markdown into OneNote page XML fragments.

OneNote's COM API keeps only simple inline HTML inside one:T CDATA. Block
structure (paragraphs, lists, tables, code blocks) has to be written as native
one:OE / one:List / one:Table elements, which is what this module produces.
"""

from __future__ import annotations

import html
import re
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from html.parser import HTMLParser

from markdown_it import MarkdownIt

from .constants import ONE


HEADING_STYLES = {
    "h1": "font-size:20.0pt;font-weight:bold",
    "h2": "font-size:16.0pt;font-weight:bold",
    "h3": "font-size:14.0pt;font-weight:bold",
    "h4": "font-size:12.0pt;font-weight:bold",
    "h5": "font-size:11.0pt;font-weight:bold",
    "h6": "font-size:10.0pt;font-weight:bold",
}

INLINE_STYLE_TAGS = {
    "code": "font-family:Consolas,'Courier New',monospace",
    "kbd": "font-family:Consolas,'Courier New',monospace",
    "del": "text-decoration:line-through",
    "mark": "background:#FFF2CC",
    "s": "text-decoration:line-through",
    "strike": "text-decoration:line-through",
}

PASSTHROUGH_INLINE_TAGS = {"a", "b", "em", "i", "span", "strong", "sub", "sup", "u"}

SAFE_ATTRS = {
    "a": {"href", "title"},
    "span": {"style"},
}

SAFE_HREF_PREFIXES = ("http://", "https://", "onenote:", "mailto:")

CODE_STYLE = INLINE_STYLE_TAGS["code"]
QUOTE_STYLE = "color:#595959;font-style:italic"
RULE_HTML = '<span style="color:#A6A6A6">' + "─" * 24 + "</span>"

VOID_TAGS = {"area", "base", "br", "col", "embed", "hr", "img", "input", "link", "meta", "source", "track", "wbr"}
DROP_TAGS = {"script", "style", "head", "title", "template", "noscript"}
CONTAINER_TAGS = {
    "address",
    "article",
    "aside",
    "body",
    "center",
    "dd",
    "details",
    "div",
    "dl",
    "dt",
    "figcaption",
    "figure",
    "footer",
    "header",
    "html",
    "li",
    "main",
    "nav",
    "section",
    "summary",
}
BLOCK_TAGS = CONTAINER_TAGS | set(HEADING_STYLES) | {"blockquote", "hr", "ol", "p", "pre", "table", "ul"}

BULLET_SYMBOLS = ("2", "3")
TASK_PREFIX = re.compile(r"^\s*\[( |x|X)\]\s+")
CHECKBOX = {False: "☐ ", True: "☑ "}

_markdown = MarkdownIt("commonmark", {"html": True}).enable(["table", "strikethrough"])


@dataclass
class Node:
    tag: str
    attrs: dict[str, str] = field(default_factory=dict)
    children: list["Node | str"] = field(default_factory=list)


@dataclass
class TableCell:
    html: str
    header: bool = False


class _TreeBuilder(HTMLParser):
    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.root = Node("root")
        self._stack = [self.root]
        self._drop_depth = 0

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        tag = tag.lower()
        if tag in DROP_TAGS:
            self._drop_depth += 1
            return
        if self._drop_depth:
            return
        node = Node(tag, {name.lower(): value for name, value in attrs if value is not None})
        self._stack[-1].children.append(node)
        if tag not in VOID_TAGS:
            self._stack.append(node)

    def handle_startendtag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        self.handle_starttag(tag, attrs)
        if tag.lower() not in VOID_TAGS and tag.lower() not in DROP_TAGS and not self._drop_depth:
            self._stack.pop()

    def handle_endtag(self, tag: str) -> None:
        tag = tag.lower()
        if tag in DROP_TAGS:
            self._drop_depth = max(0, self._drop_depth - 1)
            return
        if self._drop_depth:
            return
        for index in range(len(self._stack) - 1, 0, -1):
            if self._stack[index].tag == tag:
                del self._stack[index:]
                return

    def handle_data(self, data: str) -> None:
        if not self._drop_depth and data:
            self._stack[-1].children.append(data)


def parse_html(content: str) -> Node:
    builder = _TreeBuilder()
    builder.feed(content or "")
    builder.close()
    return builder.root


def markdown_to_html(content: str) -> str:
    """Render CommonMark (plus GFM tables and strikethrough) to HTML."""

    return _markdown.render(content or "")


def _is_block(node: Node | str) -> bool:
    return isinstance(node, Node) and node.tag in BLOCK_TAGS


def _text_content(node: Node | str) -> str:
    if isinstance(node, str):
        return node
    return "".join(_text_content(child) for child in node.children)


def _render_attrs(node: Node) -> str:
    allowed = SAFE_ATTRS.get(node.tag, set())
    parts = []
    for name, value in node.attrs.items():
        if name not in allowed:
            continue
        if name == "href" and not value.strip().lower().startswith(SAFE_HREF_PREFIXES):
            continue
        parts.append(f'{name}="{html.escape(value, quote=True)}"')
    return (" " + " ".join(parts)) if parts else ""


def render_inline(nodes: list[Node | str], *, preserve_whitespace: bool = False) -> str:
    """Render inline nodes to the inline HTML subset OneNote keeps in one:T.

    Block descendants are flattened into <br/>-separated text.
    """

    parts: list[str] = []
    for node in nodes:
        if isinstance(node, str):
            text = node if preserve_whitespace else re.sub(r"\s+", " ", node)
            parts.append(html.escape(text, quote=False))
            continue
        tag = node.tag
        if tag == "br":
            parts.append("<br/>")
        elif tag == "hr":
            parts.append("<br/>")
        elif tag == "img":
            alt = node.attrs.get("alt", "").strip()
            if alt:
                parts.append(html.escape(f"[{alt}]", quote=False))
        elif tag == "input":
            if node.attrs.get("type", "").lower() == "checkbox":
                parts.append(CHECKBOX["checked" in node.attrs])
        elif tag in HEADING_STYLES:
            inner = render_inline(node.children, preserve_whitespace=preserve_whitespace)
            parts.append(f'<br/><span style="{HEADING_STYLES[tag]}">{inner}</span><br/>')
        elif tag in INLINE_STYLE_TAGS:
            inner = render_inline(node.children, preserve_whitespace=preserve_whitespace)
            parts.append(f'<span style="{INLINE_STYLE_TAGS[tag]}">{inner}</span>')
        elif tag in PASSTHROUGH_INLINE_TAGS:
            inner = render_inline(node.children, preserve_whitespace=preserve_whitespace)
            parts.append(f"<{tag}{_render_attrs(node)}>{inner}</{tag}>")
        elif tag in BLOCK_TAGS:
            inner = render_inline(node.children, preserve_whitespace=preserve_whitespace)
            parts.append(f"<br/>{inner}<br/>")
        else:
            parts.append(render_inline(node.children, preserve_whitespace=preserve_whitespace))
    return "".join(parts)


def _tidy_breaks(fragment: str) -> str:
    fragment = re.sub(r"\s*<br/>\s*", "<br/>", fragment)
    fragment = re.sub(r"(?:<br/>){3,}", "<br/><br/>", fragment)
    fragment = re.sub(r"^(?:<br/>)+|(?:<br/>)+$", "", fragment)
    return fragment.strip()


def html_to_inline(content: str) -> str:
    """Flatten an HTML fragment to one inline fragment; blocks become <br/> breaks."""

    return _tidy_breaks(render_inline(parse_html(content).children))


def cdata(value: str) -> str:
    return "<![CDATA[" + value.replace("]]>", "]]]]><![CDATA[>") + "]]>"


def one_t(fragment_html: str) -> str:
    return f"<one:T>{cdata(fragment_html)}</one:T>"


def _styled(fragment: str, style: str | None) -> str:
    if not style or not fragment:
        return fragment
    return f'<span style="{style}">{fragment}</span>'


QUICK_STYLES = {
    "h1": 'font="Calibri Light" fontSize="20.0" fontColor="#1E4E79"',
    "h2": 'font="Calibri Light" fontSize="16.0" fontColor="#2E75B5"',
    "h3": 'font="Calibri Light" fontSize="14.0" fontColor="#1F3763"',
    "h4": 'font="Calibri Light" fontSize="12.0" fontColor="#2F5496" italic="true"',
    "h5": 'font="Calibri" fontSize="11.0" fontColor="#2E75B5" bold="true"',
    "h6": 'font="Calibri" fontSize="11.0" fontColor="#1E4E79" italic="true"',
    "blockquote": 'font="Calibri" fontSize="11.0" fontColor="#595959" italic="true"',
    "code": 'font="Consolas" fontSize="10.0" fontColor="automatic"',
}
TODO_TAG_DEF = 'type="0" symbol="3" fontColor="automatic" highlightColor="none" name="To Do"'


class StyleSet:
    """Native OneNote styles (QuickStyleDef) and the To Do tag (TagDef) used by one page update.

    Indexes only need to be consistent inside the submitted XML; OneNote maps them
    onto the page's own definitions by name. When editing an existing page, seed
    the set with that page's definitions so existing paragraphs keep their indexes.
    """

    def __init__(self, existing_defs: list[ET.Element] | None = None) -> None:
        self.existing = list(existing_defs or [])
        self.quick: dict[str, int] = {}
        self.new_quick: list[str] = []
        self.todo_index: int | None = None
        self.new_todo = False
        self._next_quick = 0
        self._next_tag = 0
        for node in self.existing:
            index = int(node.attrib.get("index", "0"))
            if node.tag == ONE + "QuickStyleDef":
                self.quick.setdefault(node.attrib.get("name", ""), index)
                self._next_quick = max(self._next_quick, index + 1)
            elif node.tag == ONE + "TagDef":
                if self.todo_index is None and node.attrib.get("type") == "0" and node.attrib.get("symbol") == "3":
                    self.todo_index = index
                self._next_tag = max(self._next_tag, index + 1)

    def quick_index(self, name: str) -> int:
        if name not in self.quick:
            self.quick[name] = self._next_quick
            self._next_quick += 1
            self.new_quick.append(name)
        return self.quick[name]

    def todo_tag(self) -> int:
        if self.todo_index is None:
            self.todo_index = self._next_tag
            self._next_tag += 1
            self.new_todo = True
        return self.todo_index

    def defs_xml(self) -> str:
        """TagDefs, then QuickStyleDefs (the schema order); existing definitions are copied unchanged."""

        tags = [ET.tostring(node, encoding="unicode") for node in self.existing if node.tag == ONE + "TagDef"]
        if self.new_todo:
            tags.append(f'<one:TagDef index="{self.todo_index}" {TODO_TAG_DEF}/>')
        quick = [ET.tostring(node, encoding="unicode") for node in self.existing if node.tag == ONE + "QuickStyleDef"]
        quick.extend(
            f'<one:QuickStyleDef index="{self.quick[name]}" name="{name}" highlightColor="automatic" '
            f'spaceBefore="0.0" spaceAfter="0.0" {QUICK_STYLES[name]}/>'
            for name in self.new_quick
        )
        return "".join(tags + quick)


def _oe(
    fragment: str,
    *,
    list_xml: str = "",
    children: list[str] | None = None,
    quick: str | None = None,
    styles: StyleSet | None = None,
    todo: bool | None = None,
) -> str:
    children_xml = f"<one:OEChildren>{''.join(children)}</one:OEChildren>" if children else ""
    list_part = f"<one:List>{list_xml}</one:List>" if list_xml else ""
    attrs = ""
    if quick and styles is not None:
        attrs = f' quickStyleIndex="{styles.quick_index(quick)}"'
    elif quick:
        fallback = {"blockquote": QUOTE_STYLE, "code": CODE_STYLE}.get(quick, HEADING_STYLES.get(quick))
        fragment = _styled(fragment, fallback)
    tag_part = ""
    if todo is not None and styles is not None:
        tag_part = f'<one:Tag index="{styles.todo_tag()}" completed="{"true" if todo else "false"}"/>'
    elif todo is not None:
        fragment = CHECKBOX[todo] + fragment
    return f"<one:OE{attrs}>{tag_part}{list_part}{one_t(fragment)}{children_xml}</one:OE>"


@dataclass
class _Context:
    quick: str | None = None
    list_depth: int = 0
    styles: StyleSet | None = None

    def nested(self, **changes: object) -> "_Context":
        values = {"quick": self.quick, "list_depth": self.list_depth, "styles": self.styles}
        values.update(changes)
        return _Context(**values)  # type: ignore[arg-type]

    def oe(self, fragment: str, **kwargs: object) -> str:
        kwargs.setdefault("quick", self.quick)
        return _oe(fragment, styles=self.styles, **kwargs)  # type: ignore[arg-type]


def _paragraphs(fragment: str, context: _Context, style: str = "") -> list[str]:
    fragment = _tidy_breaks(fragment)
    if not fragment:
        return []
    return [context.oe(_styled(part, style)) for part in fragment.split("<br/>")]


def _list_marker(ordered: bool, number: int, depth: int) -> str:
    if ordered:
        return f'<one:Number numberSequence="0" numberFormat="##." text="{number}."/>'
    return f'<one:Bullet bullet="{BULLET_SYMBOLS[depth % len(BULLET_SYMBOLS)]}"/>'


def _list_items(node: Node) -> list[Node]:
    return [child for child in node.children if isinstance(child, Node) and child.tag == "li"]


def _render_list(node: Node, context: _Context) -> list[str]:
    ordered = node.tag == "ol"
    try:
        number = int(node.attrs.get("start", "1"))
    except ValueError:
        number = 1
    items = []
    for item in _list_items(node):
        items.append(_render_list_item(item, ordered, number, context))
        number += 1
    return items


def _render_list_item(item: Node, ordered: bool, number: int, context: _Context) -> str:
    children = list(item.children)
    while children and isinstance(children[0], str) and not children[0].strip():
        children.pop(0)

    lead: list[Node | str] = []
    if children and isinstance(children[0], Node) and children[0].tag == "p":
        lead = list(children.pop(0).children)
    else:
        while children and not _is_block(children[0]):
            lead.append(children.pop(0))

    lead_html = _tidy_breaks(render_inline(lead))
    plain_lead = html.unescape(re.sub(r"<[^>]+>", "", lead_html))
    todo: bool | None = None
    task = TASK_PREFIX.match(plain_lead)
    if task:
        lead_html = re.sub(r"\[( |x|X)\]\s+", "", lead_html, count=1)
        todo = task.group(1) != " "
    else:
        for checked, mark in CHECKBOX.items():
            if lead_html.startswith(mark):
                lead_html = lead_html[len(mark) :]
                todo = checked

    first, *rest = lead_html.split("<br/>") if lead_html else [""]
    child_context = context.nested(list_depth=context.list_depth + 1)
    child_oes = [context.oe(part) for part in rest if part]
    child_oes.extend(_render_blocks(children, child_context))
    return context.oe(
        first,
        list_xml="" if todo is not None else _list_marker(ordered, number, context.list_depth),
        children=child_oes,
        todo=todo,
    )


def _render_pre(node: Node, context: _Context) -> list[str]:
    text = _text_content(node).replace("\r\n", "\n").replace("\r", "\n")
    if text.endswith("\n"):
        text = text[:-1]
    paragraphs = []
    for line in text.split("\n"):
        stripped = line.lstrip(" ")
        indent = "&#160;" * (len(line) - len(stripped))
        body = indent + html.escape(stripped.replace("\t", "    "), quote=False)
        paragraphs.append(context.oe(body, quick="code"))
    return paragraphs


def _table_rows(node: Node) -> list[list[TableCell]]:
    rows: list[list[TableCell]] = []

    def visit(current: Node) -> None:
        for child in current.children:
            if not isinstance(child, Node) or child.tag == "table":
                continue
            if child.tag == "tr":
                cells = [
                    TableCell(html=_tidy_breaks(render_inline(cell.children)), header=cell.tag == "th")
                    for cell in child.children
                    if isinstance(cell, Node) and cell.tag in {"td", "th"}
                ]
                if cells:
                    rows.append(cells)
            else:
                visit(child)

    visit(node)
    return rows


def _one_table_cell(cell: TableCell) -> str:
    shading = ' shadingColor="#D9EAF7"' if cell.header else ""
    font_size = "10.5pt" if cell.header else "10.0pt"
    style_attr = f' style="font-family:\'Microsoft YaHei\';font-size:{font_size}"'
    cell_html = cell.html
    if cell.header:
        cell_html = f"<span style='font-weight:bold'>{cell_html}</span>"
    return (
        f"<one:Cell{shading}>"
        "<one:OEChildren>"
        f'<one:OE alignment="left"{style_attr}>{one_t(cell_html)}</one:OE>'
        "</one:OEChildren>"
        "</one:Cell>"
    )


TABLE_MAX_WIDTH = 960.0
COLUMN_MIN_WIDTH = 48.0
COLUMN_MAX_WIDTH = 320.0
CELL_PADDING = 14.0
_TAGS = re.compile(r"<[^>]+>")


def _text_width(cell: TableCell) -> float:
    """Rough width in points of a cell's longest line at the 10-10.5 pt table font."""

    size = 10.5 if cell.header else 10.0
    lines = html.unescape(_TAGS.sub("\n", cell.html.replace("<br/>", "\n"))).split("\n")
    return max((sum(size if ord(ch) > 0x2E7F else size * 0.56 for ch in line) for line in lines), default=0.0)


def _table_column_widths(rows: list[list[TableCell]]) -> list[float]:
    """Each column as wide as its longest cell, within limits; the table is shrunk to fit a normal page."""

    column_count = max((len(row) for row in rows), default=0)
    if column_count <= 0:
        return []
    widths = [COLUMN_MIN_WIDTH] * column_count
    for row in rows:
        for index, cell in enumerate(row[:column_count]):
            widths[index] = max(widths[index], min(COLUMN_MAX_WIDTH, _text_width(cell) + CELL_PADDING))
    total = sum(widths)
    if total > TABLE_MAX_WIDTH:
        widths = [max(COLUMN_MIN_WIDTH, width * TABLE_MAX_WIDTH / total) for width in widths]
    return widths


def one_table(rows: list[list[TableCell]]) -> str:
    widths = _table_column_widths(rows)
    if not widths:
        return ""
    column_xml = "".join(
        f'<one:Column index="{index}" width="{width:.1f}" isLocked="true"/>'
        for index, width in enumerate(widths)
    )
    column_count = len(widths)
    row_xml = []
    for row in rows:
        padded = row + [TableCell(html="")] * (column_count - len(row))
        row_xml.append("<one:Row>" + "".join(_one_table_cell(cell) for cell in padded[:column_count]) + "</one:Row>")
    return (
        '<one:OE alignment="left"><one:Table bordersVisible="true" hasHeaderRow="false">'
        f"<one:Columns>{column_xml}</one:Columns>"
        f"{''.join(row_xml)}"
        "</one:Table></one:OE>"
    )


def _render_block(node: Node, context: _Context) -> list[str]:
    tag = node.tag
    if tag in HEADING_STYLES:
        inner = _tidy_breaks(render_inline(node.children))
        return [context.oe(part, quick=tag) for part in inner.split("<br/>") if part]
    if tag == "p":
        return _paragraphs(render_inline(node.children), context, html.escape(node.attrs.get("style", ""), quote=True))
    if tag in {"ul", "ol"}:
        return _render_list(node, context)
    if tag == "pre":
        return _render_pre(node, context)
    if tag == "table":
        rows = _table_rows(node)
        return [one_table(rows)] if rows else []
    if tag == "blockquote":
        return _render_blocks(node.children, context.nested(quick="blockquote"))
    if tag == "hr":
        return [_oe(RULE_HTML)]
    return _render_blocks(node.children, context)


def _render_blocks(nodes: list[Node | str], context: _Context) -> list[str]:
    output: list[str] = []
    run: list[Node | str] = []

    def flush() -> None:
        if run:
            output.extend(_paragraphs(render_inline(run), context))
            run.clear()

    for node in nodes:
        if _is_block(node):
            flush()
            output.extend(_render_block(node, context))  # type: ignore[arg-type]
        else:
            run.append(node)
    flush()
    return output


def html_to_oe_xml(content: str, styles: StyleSet | None = None) -> str:
    """Convert an HTML fragment to a sequence of one:OE elements.

    With a StyleSet, headings, quotes, code and to-do items use native OneNote
    styles and tags; the caller must emit styles.defs_xml() in the page update.
    Without one, they fall back to inline formatting.
    """

    return "".join(_render_blocks(parse_html(content).children, _Context(styles=styles)))


def plain_to_oe_xml(content: str) -> str:
    escaped = html.escape(content, quote=False)
    normalized = escaped.replace("\r\n", "\n").replace("\r", "\n")
    return "".join(_oe(line) for line in normalized.split("\n"))
