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
from markdown_it.rules_inline.state_inline import Scanned, StateInline

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

SAFE_HREF_PREFIXES = ("http://", "https://", "onenote:", "mailto:", "file:")

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
# Every other tag name is text that only looks like a tag, e.g. List<String> or <T> in prose; it is kept literally.
KNOWN_TAGS = (
    BLOCK_TAGS
    | VOID_TAGS
    | DROP_TAGS
    | set(INLINE_STYLE_TAGS)
    | PASSTHROUGH_INLINE_TAGS
    | {
        "abbr", "acronym", "audio", "bdi", "bdo", "big", "button", "canvas", "caption", "cite", "colgroup", "data",
        "del", "dfn", "dialog", "fieldset", "font", "form", "iframe", "ins", "label", "legend", "map", "menu", "meter",
        "nobr", "object", "optgroup", "option", "output", "picture", "progress", "q", "rp", "rt", "ruby", "samp",
        "select", "small", "svg", "tbody", "td", "textarea", "tfoot", "th", "thead", "time", "tr", "tt", "var", "video",
    }
)

BULLET_SYMBOLS = ("2", "3")
TASK_PREFIX = re.compile(r"^\s*\[( |x|X)\]\s+")
CHECKBOX = {False: "☐ ", True: "☑ "}

_markdown = MarkdownIt("commonmark", {"html": True}).enable(["table", "strikethrough"])


def _is_cjk(char: str) -> bool:
    code = ord(char)
    return 0x2E80 <= code <= 0x9FFF or 0xAC00 <= code <= 0xD7AF or 0xF900 <= code <= 0xFAFF or 0xFE30 <= code <= 0xFE4F or 0xFF00 <= code <= 0xFFEF


def _cjk_friendly_scan_delims(self: StateInline, start: int, canSplitWord: bool) -> Scanned:
    """CommonMark delimiter rules, relaxed next to CJK characters.

    CommonMark does not let `**` close after punctuation when a letter follows, so `**【重点】**内容` stays
    literal. Chinese text has no spaces around emphasis, so a CJK neighbour counts like whitespace here.
    """

    scanned = _scan_delims(self, start, canSplitWord)
    last = self.src[start - 1] if start > 0 else " "
    end = start + scanned.length
    after = self.src[end] if end < self.posMax else " "
    if not (_is_cjk(last) or _is_cjk(after)) or self.src[start] != "*":
        return scanned
    can_open = not after.isspace() and (scanned.can_open or _is_cjk(last) or _is_cjk(after))
    can_close = not last.isspace() and (scanned.can_close or _is_cjk(after) or _is_cjk(last))
    return Scanned(can_open, can_close, scanned.length)


_scan_delims = StateInline.scanDelims
StateInline.scanDelims = _cjk_friendly_scan_delims  # type: ignore[method-assign]


@dataclass
class Node:
    tag: str
    attrs: dict[str, str] = field(default_factory=dict)
    children: list["Node | str"] = field(default_factory=list)


@dataclass
class TableCell:
    html: str
    header: bool = False
    align: str = "left"


class _TreeBuilder(HTMLParser):
    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.root = Node("root")
        self._stack = [self.root]
        self._drop_depth = 0

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        tag = tag.lower()
        if tag not in KNOWN_TAGS and ":" not in tag:
            self.handle_data(self.get_starttag_text() or f"<{tag}>")
            return
        if tag in DROP_TAGS:
            self._drop_depth += 1
            return
        if self._drop_depth:
            return
        node = Node(tag, {name.lower(): value or "" for name, value in attrs})
        self._stack[-1].children.append(node)
        if tag not in VOID_TAGS:
            self._stack.append(node)

    def handle_startendtag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        if tag.lower() not in KNOWN_TAGS and ":" not in tag:
            self.handle_data(self.get_starttag_text() or f"<{tag}/>")
            return
        self.handle_starttag(tag, attrs)
        if tag.lower() not in VOID_TAGS and tag.lower() not in DROP_TAGS and not self._drop_depth:
            self._stack.pop()

    def handle_endtag(self, tag: str) -> None:
        tag = tag.lower()
        if tag not in KNOWN_TAGS and ":" not in tag:
            self.handle_data(f"</{tag}>")
            return
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

    # A paragraph's leading no-break spaces are indentation (get_page writes them so); the parser would trim them.
    content = re.sub(r"(?m)^( {0,3})(\xa0+)", lambda m: m.group(1) + "&#160;" * len(m.group(2)), content or "")
    return _markdown.render(content)


def markdown_inline_to_html(content: str) -> str:
    """Render only inline Markdown (emphasis, code, links), so '2026.', '# x' or '- x' stay plain text."""

    return _markdown.renderInline(content or "")


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


def _collapse(text: str) -> str:
    """HTML whitespace rules for source line breaks, but runs of spaces typed on one line are kept.

    A line break and the indentation around it is source formatting and becomes one space. Several spaces in a
    row are meant (aligned comments, ASCII diagrams), so they become no-break spaces, which OneNote keeps.
    """

    text = re.sub(r"[ \t\r\f\v]*\n[ \t\r\n\f\v]*", " ", text.replace("\t", "    "))
    return re.sub(r" {2,}", lambda run: " " + "\xa0" * (len(run.group()) - 1), text)


def render_inline(nodes: list[Node | str], *, preserve_whitespace: bool = False) -> str:
    """Render inline nodes to the inline HTML subset OneNote keeps in one:T.

    Block descendants are flattened into <br/>-separated text.
    """

    parts: list[str] = []
    for node in nodes:
        if isinstance(node, str):
            text = node if preserve_whitespace else _collapse(node)
            parts.append(html.escape(text, quote=False))
            continue
        tag = node.tag
        if tag == "br":
            parts.append("<br/>")
        elif tag == "hr":
            parts.append("<br/>")
        elif tag == "img":
            alt = " ".join(node.attrs.get("alt", "").split())
            source = node.attrs.get("src", "").strip()
            label = alt or (source if not source.startswith("data:") else "")
            parts.append(html.escape(f"[Image: {label}]" if label else "[Image]", quote=False))
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
            href = node.attrs.get("href", "").strip()
            if tag == "a" and href and not href.lower().startswith(SAFE_HREF_PREFIXES):
                # OneNote cannot follow a relative link; keep where it pointed as visible text.
                scheme = re.match(r"^([a-z][a-z0-9+.-]*):", href, re.I)
                target = "" if scheme and scheme.group(1).lower() in {"javascript", "data", "vbscript"} else href
                parts.append(inner + (html.escape(f" ({target})", quote=False) if target and target != _text_content(node).strip() else ""))
                continue
            parts.append(f"<{tag}{_render_attrs(node)}>{inner}</{tag}>")
        elif tag in BLOCK_TAGS:
            inner = render_inline(node.children, preserve_whitespace=preserve_whitespace)
            parts.append(f"<br/>{inner}<br/>")
        else:
            parts.append(render_inline(node.children, preserve_whitespace=preserve_whitespace))
    return "".join(parts)


def _tidy_breaks(fragment: str) -> str:
    # Only ordinary whitespace is trimmed; no-break spaces are deliberate indentation.
    fragment = re.sub(r"[ \t\r\n]*<br/>[ \t\r\n]*", "<br/>", fragment)
    fragment = re.sub(r"(?:<br/>){3,}", "<br/><br/>", fragment)
    fragment = re.sub(r"^(?:<br/>)+|(?:<br/>)+$", "", fragment)
    return fragment.strip(" \t\r\n")


def html_to_inline(content: str) -> str:
    """Flatten an HTML fragment to one inline fragment; blocks become <br/> breaks."""

    return _tidy_breaks(render_inline(parse_html(content).children))


_XML_INVALID = re.compile("[\x00-\x08\x0b\x0c\x0e-\x1f￾￿]")


def cdata(value: str) -> str:
    value = _XML_INVALID.sub("", value)
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


def _list_marker(ordered: bool, number: int, depth: int, restart: bool = False) -> str:
    if ordered:
        # OneNote renumbers lists itself and continues a list after a paragraph break; restartNumberingAt on
        # the first item keeps a list's own start (5. ...) and makes a new list start over.
        restart_attr = f' restartNumberingAt="{number}"' if restart else ""
        return f'<one:Number numberSequence="0" numberFormat="##."{restart_attr} text="{number}."/>'
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
    for position, item in enumerate(_list_items(node)):
        items.append(_render_list_item(item, ordered, number, context, restart=position == 0))
        number += 1
    return items


def _render_list_item(item: Node, ordered: bool, number: int, context: _Context, restart: bool = False) -> str:
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
        # A to-do in a bulleted list shows only its box; a numbered to-do keeps its number too.
        list_xml="" if todo is not None and not ordered else _list_marker(ordered, number, context.list_depth, restart),
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


def _alignment(cell: Node) -> str:
    match = re.search(r"text-align\s*:\s*(left|center|right)", cell.attrs.get("style", ""), re.I)
    return (match.group(1) if match else cell.attrs.get("align", "left")).lower()


def _table_rows(node: Node) -> list[list[TableCell]]:
    rows: list[list[TableCell]] = []

    def visit(current: Node) -> None:
        for child in current.children:
            if not isinstance(child, Node) or child.tag == "table":
                continue
            if child.tag == "tr":
                cells = [
                    TableCell(
                        html=_tidy_breaks(render_inline(cell.children)),
                        header=cell.tag == "th",
                        align=_alignment(cell),
                    )
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
    alignment = cell.align if cell.align in {"left", "center", "right"} else "left"
    # OneNote drops <br> inside a cell's text, so each line of a cell is its own paragraph.
    paragraphs = []
    for line in (cell.html.split("<br/>") if cell.html else [""]):
        if cell.header and line:
            line = f"<span style='font-weight:bold'>{line}</span>"
        paragraphs.append(f'<one:OE alignment="{alignment}"{style_attr}>{one_t(line)}</one:OE>')
    return f"<one:Cell{shading}><one:OEChildren>{''.join(paragraphs)}</one:OEChildren></one:Cell>"


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
        return _render_quote(node, context.nested(quick="blockquote"))
    if tag == "hr":
        return [_oe(RULE_HTML)]
    return _render_blocks(node.children, context)


QUOTE_NESTED_TAGS = {"blockquote", "pre"}


def _with_children(oe_xml: str, children: list[str]) -> str:
    """Add paragraphs as indented children of an already rendered one:OE."""

    closing = "</one:OEChildren></one:OE>"
    if oe_xml.endswith(closing):
        return oe_xml[: -len(closing)] + "".join(children) + closing
    return oe_xml[: -len("</one:OE>")] + "<one:OEChildren>" + "".join(children) + closing


def _render_quote(node: Node, context: _Context) -> list[str]:
    """A quote; a quote or code block inside it is indented under the paragraph before it, so it stays inside."""

    output: list[str] = []
    run: list[Node | str] = []
    for child in node.children:
        if isinstance(child, Node) and child.tag in QUOTE_NESTED_TAGS and output + run:
            output.extend(_render_blocks(run, context))
            run = []
            nested = _render_block(child, context)
            if output and nested:
                output[-1] = _with_children(output[-1], nested)
            else:
                output.extend(nested)
        else:
            run.append(child)
    output.extend(_render_blocks(run, context))
    return output


def _render_blocks(nodes: list[Node | str], context: _Context) -> list[str]:
    output: list[str] = []
    run: list[Node | str] = []
    previous = ""

    def flush() -> None:
        if run:
            output.extend(_paragraphs(render_inline(run), context))
            run.clear()

    for node in nodes:
        if _is_block(node):
            flush()
            if node.tag == "pre" and previous == "pre":  # type: ignore[union-attr]
                # Two code blocks in a row would read back as one; an empty line keeps them apart.
                output.append(context.oe(""))
            output.extend(_render_block(node, context))  # type: ignore[arg-type]
            previous = node.tag  # type: ignore[union-attr]
        else:
            run.append(node)
            if isinstance(node, Node) or node.strip():
                previous = ""
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
