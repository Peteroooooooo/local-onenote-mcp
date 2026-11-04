"""Render OneNote page XML as readable plain text or Markdown."""

from __future__ import annotations

import re
import xml.etree.ElementTree as ET
from html.parser import HTMLParser
from typing import Literal

from .constants import ONE_NS


ReadFormat = Literal["text", "markdown"]
_NS = f"{{{ONE_NS}}}"
INK_LABEL = "[Handwriting: no recognised text; use view_page to see it]"


def _local(tag: str) -> str:
    return tag.rsplit("}", 1)[-1] if "}" in tag else tag


def _child(node: ET.Element, name: str) -> ET.Element | None:
    return node.find(_NS + name)


def _children(node: ET.Element, name: str) -> list[ET.Element]:
    return node.findall(_NS + name)


def _style_marks(tag: str, style: str) -> set[str]:
    style = re.sub(r"\s+", "", style).lower()
    marks = set()
    if tag in {"b", "strong"} or "font-weight:bold" in style:
        marks.add("**")
    if tag in {"i", "em"} or "font-style:italic" in style:
        marks.add("*")
    if tag in {"s", "del", "strike"} or "line-through" in style:
        marks.add("~~")
    if _is_mono(style):
        marks.add("`")
    return marks


def _is_mono(style: str) -> bool:
    family = re.search(r"font-family:([^;]*)", re.sub(r"\s+", "", style).lower())
    return bool(family) and ("consolas" in family.group(1) or "courier" in family.group(1))


_ESCAPES = (
    (re.compile(r"\\(?=[!-/:-@\[-`{-~])"), r"\\\\"),
    (re.compile(r"([*`])"), r"\\\1"),
    (re.compile(r"(?<![0-9A-Za-z])_|_(?![0-9A-Za-z])"), r"\\_"),
    (re.compile(r"~(?=~)"), r"\\~"),
    (re.compile(r"<(?=[A-Za-z/!?])"), r"\\<"),
    (re.compile(r"&(?=#?\w+;)"), r"\\&"),
    (re.compile(r"\[(?=[^\]]*\]\s*[(\[:])"), r"\\["),
)
_LINE_START = re.compile(r"^(?:#{1,6}(?=\s|$)|[-+](?=\s|$)|>|=+\s*$|-{2,}\s*$|\|)")
_NUMBERED_START = re.compile(r"^(\d{1,9})([.)])(?=\s|$)")


def _escape(text: str) -> str:
    """Backslash-escape text so Markdown shows it literally."""

    for pattern, replacement in _ESCAPES:
        text = pattern.sub(replacement, text)
    return text


def _escape_line_start(line: str) -> str:
    """Escape what would turn a line into a heading, list item, quote, rule or table row."""

    numbered = _NUMBERED_START.match(line)
    if numbered:
        return f"{numbered.group(1)}\\{line[numbered.end(1):]}"
    return "\\" + line if _LINE_START.match(line) else line


def _code_span(text: str) -> str:
    longest = max((len(run) for run in re.findall(r"`+", text)), default=0)
    fence = "`" * (longest + 1)
    pad = " " if text.startswith("`") or text.endswith("`") else ""
    return f"{fence}{pad}{text}{pad}{fence}"


def _link_target(url: str) -> str:
    return f"<{url}>" if re.search(r"[\s()<>]", url) else url


class _InlineReader(HTMLParser):
    """Turn a one:T inline HTML fragment into text, optionally with Markdown marks.

    OneNote splits text into runs whenever the font changes (Chinese and Latin text get separate runs), so the
    Markdown output merges neighbouring runs with the same formatting before adding marks. base_style is the
    style of the paragraph itself (OneNote moves formatting that covers a whole paragraph onto the one:OE).
    """

    def __init__(self, markdown: bool, ignore: str = "", base_style: str = "") -> None:
        super().__init__(convert_charrefs=True)
        self.markdown = markdown
        self.ignore = ignore
        self.segments: list[tuple[str, tuple[str, ...], str]] = []  # (text, marks, link)
        self._marks: list[set[str]] = [_style_marks("", base_style)] if base_style else []
        self._base_depth = len(self._marks)
        self._links: list[str] = []

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        tag = tag.lower()
        if tag == "br":
            self.handle_data("\n", raw=False)
            return
        attributes = {name.lower(): value or "" for name, value in attrs}
        if tag == "a":
            self._links.append(attributes.get("href", ""))
            return
        self._marks.append(_style_marks(tag, attributes.get("style", "")))

    def handle_endtag(self, tag: str) -> None:
        tag = tag.lower()
        if tag == "a":
            if self._links:
                self._links.pop()
        elif tag != "br" and len(self._marks) > self._base_depth:
            self._marks.pop()

    def handle_data(self, data: str, raw: bool = True) -> None:
        if raw:
            # A raw line break in the CDATA is only source formatting (OneNote wraps long lines); <br> is the break.
            data = re.sub(r" ?\r?\n ?", " ", data)
        held = set().union(*self._marks)
        marks = tuple(mark for mark in ("**", "*", "~~", "`") if mark in held and mark not in self.ignore)
        link = self._links[-1] if self._links else ""
        data = data.replace("\xa0", " ")
        if self.segments and (self.segments[-1][1:] == (marks, link) or (not data.strip() and not link)):
            previous, marks, link = self.segments.pop()
            data = previous + data
        self.segments.append((data, marks, link))

    def result(self) -> str:
        if not self.markdown:
            return "".join(text for text, _, _ in self.segments).replace("\x00", "")
        lines: list[list[tuple[str, tuple[str, ...], str]]] = [[]]
        for text, marks, link in self.segments:
            for number, piece in enumerate(text.split("\n")):
                if number:
                    lines.append([])
                lines[-1].append((piece, marks, link))
        return "\n".join(self._marked_line(line) for line in lines).replace("\x00", "")

    @staticmethod
    def _marked_line(segments: list[tuple[str, tuple[str, ...], str]]) -> str:
        """One line with emphasis opened and closed only where it changes, so nested marks stay valid."""

        out = ""
        open_marks: list[str] = []

        def close(count: int) -> None:
            nonlocal out
            body = out.rstrip()
            space = out[len(body) :]
            closing = "".join(reversed(open_marks[len(open_marks) - count :]))
            del open_marks[len(open_marks) - count :]
            out = body + closing + space

        for text, marks, link in segments:
            core = text.strip()
            if not core:
                out += text
                continue
            lead, trail = text[: len(text) - len(text.lstrip())], text[len(text.rstrip()) :]
            wanted = [mark for mark in ("**", "*", "~~") if mark in marks]
            keep = 0
            while keep < len(open_marks) and open_marks[keep] in wanted:
                keep += 1
            if keep < len(open_marks):
                close(len(open_marks) - keep)
            opening = [mark for mark in wanted if mark not in open_marks]
            open_marks.extend(opening)
            body = _code_span(core) if "`" in marks else _escape(core)
            if link:
                body = f"[{body}]({_link_target(link)})"
            out += lead + "".join(opening) + body + trail
        if open_marks:
            close(len(open_marks))
        return out


def inline_text(fragment: str | None, *, markdown: bool = False, ignore: str = "", base_style: str = "") -> str:
    fragment = re.sub(r"(<br\s*/?>)[ \t]*\r?\n", r"\1", fragment or "", flags=re.IGNORECASE)
    parser = _InlineReader(markdown, ignore, base_style)
    parser.feed(fragment)
    parser.close()
    return parser.result()


class PageRenderer:
    def __init__(self, root: ET.Element, fmt: ReadFormat, include_ocr: bool) -> None:
        self.root = root
        self.markdown = fmt == "markdown"
        self.include_ocr = include_ocr
        self.tag_defs = {
            node.attrib.get("index", ""): node.attrib.get("name", "")
            for node in root.iter(_NS + "TagDef")
        }
        self.quick_styles = {
            node.attrib.get("index", ""): node.attrib.get("name", "")
            for node in root.iter(_NS + "QuickStyleDef")
        }

    def render(self, include_title: bool = True) -> str:
        blocks: list[str] = []
        title = self.title()
        if include_title and title:
            blocks.append(f"# {title}" if self.markdown else title)
        for node in self.root:
            kind = _local(node.tag)
            if kind == "Outline":
                lines = self._oe_children(_child(node, "OEChildren"), "")
                text = "\n".join(lines).strip("\n")
                if text:
                    blocks.append(text)
            elif kind == "Image":
                blocks.append(self._image(node))
            elif kind == "InkDrawing":
                blocks.append(INK_LABEL)
            elif kind in {"InsertedFile", "MediaFile"}:
                blocks.append(self._file(node, kind))
        return "\n\n".join(block for block in blocks if block).strip()

    def title(self) -> str | None:
        title = _child(self.root, "Title")
        if title is None:
            return None
        for node in title.iter(_NS + "T"):
            value = inline_text(node.text).strip()
            if value:
                return value
        return None

    def _oe_children(self, container: ET.Element | None, indent: str) -> list[str]:
        """Lines for the paragraphs in one container.

        In Markdown, paragraphs are separated by a blank line (a single newline would join them), items of
        one list are not, and consecutive code paragraphs form one fenced block.
        """

        if container is None:
            return []
        items: list[tuple[str, list[str]]] = []
        code: list[ET.Element] = []
        for oe in _children(container, "OE"):
            if self.markdown and self._is_code(oe):
                code.append(oe)
                continue
            if code:
                items.append(("code", self._fence(code, indent)))
                code = []
            lines = self._oe(oe, indent)
            if not any(line.strip() for line in lines):
                kind = "blank"
            elif self._list_prefix(oe) or self._tag_prefix(oe).startswith("- "):
                # A bulleted and a numbered list next to each other are two lists; keep a blank line between them.
                list_node = _child(oe, "List")
                kind = "olist" if list_node is not None and _child(list_node, "Number") is not None else "ulist"
            elif self._style_prefix(oe) == "> ":
                kind = "quote"
            else:
                kind = "para"
            items.append((kind, lines))
        if code:
            items.append(("code", self._fence(code, indent)))
        if not self.markdown:
            return [line for _, lines in items for line in lines]

        out: list[str] = []
        previous = ""
        for kind, lines in items:
            if kind == "blank":
                previous = "blank"
                continue
            if previous == kind == "quote":
                out.append(indent + ">")
            elif previous and not (previous == kind and kind in {"olist", "ulist"}):
                out.append("")
            out.extend(lines)
            previous = kind
        return out

    def _is_code(self, oe: ET.Element) -> bool:
        if self._style(oe) == "code":
            return True
        return _child(oe, "List") is None and _is_mono(oe.attrib.get("style", "")) and _child(oe, "T") is not None

    def _fence(self, oes: list[ET.Element], indent: str) -> list[str]:
        code_lines: list[str] = []

        def collect(oe: ET.Element, depth: int) -> None:
            text = "".join(inline_text(node.text) for node in _children(oe, "T"))
            code_lines.extend("    " * depth + line for line in text.split("\n"))
            children = _child(oe, "OEChildren")
            for child in _children(children, "OE") if children is not None else []:
                collect(child, depth + 1)

        for oe in oes:
            collect(oe, 0)
        longest = max((len(run) for line in code_lines for run in re.findall(r"`{3,}", line)), default=0)
        fence = "`" * max(3, longest + 1)
        return [indent + fence, *(indent + line if line.strip() else "" for line in code_lines), indent + fence]

    def _style(self, oe: ET.Element) -> str:
        return self.quick_styles.get(oe.attrib.get("quickStyleIndex", ""), "")

    def _style_prefix(self, oe: ET.Element) -> str:
        if not self.markdown:
            return ""
        style = self._style(oe)
        if re.fullmatch(r"h[1-6]", style):
            return "#" * int(style[1]) + " "
        if style in {"quote", "blockquote", "cite"}:
            return "> "
        return ""

    def _oe(self, oe: ET.Element, indent: str) -> list[str]:
        style_prefix, list_prefix, tag_prefix = self._style_prefix(oe), self._list_prefix(oe), self._tag_prefix(oe)
        prefix = style_prefix + list_prefix + tag_prefix
        marker = list_prefix or ("- " if tag_prefix.startswith("- ") else "")
        if not self.markdown:
            child_indent = indent + "  "
        else:
            # Markdown nests only under list items; indenting plain paragraphs would turn them into code.
            child_indent = indent + " " * len(marker) if marker else indent
        quote = style_prefix == "> "
        heading = style_prefix.startswith("#")
        ignore = ("**", "*") if heading else (("*",) if quote else ())
        base_style = oe.attrib.get("style", "")
        lines: list[str] = []
        body: list[str] = []
        for node in oe:
            kind = _local(node.tag)
            if kind == "T":
                body.append(inline_text(node.text, markdown=self.markdown, ignore=ignore, base_style=base_style))
            elif kind == "Table":
                lines.extend(["", *self._indent(self._table(node), indent), ""])
            elif kind == "Image":
                lines.extend(["", *self._indent(self._image(node).split("\n"), indent), ""])
            elif kind == "InkWord" and node.attrib.get("recognizedText"):
                body.append(node.attrib["recognizedText"])
            elif kind in {"InkDrawing", "InkWord"}:
                body.append(INK_LABEL)
            elif kind in {"InsertedFile", "MediaFile"}:
                lines.extend(self._indent([self._file(node, kind)], indent))
        text = "".join(body)
        if text.strip() or prefix:
            text_lines = text.split("\n")
            if self.markdown:
                text_lines = [line.strip() for line in text_lines]
                text_lines = [line for line in text_lines if line] or [""]
                text_lines = [_escape_line_start(line) for line in text_lines]
                if not heading:
                    # A <br> inside a paragraph is a hard line break.
                    text_lines = [line + "\\" for line in text_lines[:-1]] + text_lines[-1:]
            first, *rest = text_lines
            out = [f"{indent}{prefix}{first}".rstrip()]
            continuation = indent + ("> " if quote else " " * len(prefix))
            out.extend(f"{continuation}{line}".rstrip() for line in rest)
            lines = out + lines
        elif not lines and not self._has_children(oe):
            lines.append("")
        children = self._oe_children(_child(oe, "OEChildren"), child_indent)
        if self.markdown and quote and children:
            # Paragraphs indented under a quote (a nested quote, code) stay inside it.
            children = [indent + ">" + (" " + line[len(indent) :] if line.strip() else "") for line in ["", *children]]
        elif self.markdown and children and lines and lines[-1].strip():
            first_child = _children(_child(oe, "OEChildren"), "OE")[0]
            if not (self._list_prefix(first_child) or self._tag_prefix(first_child).startswith("- ")):
                lines.append("")
        lines.extend(children)
        return lines

    @staticmethod
    def _has_children(oe: ET.Element) -> bool:
        return _child(oe, "OEChildren") is not None

    @staticmethod
    def _indent(lines: list[str], indent: str) -> list[str]:
        return [f"{indent}{line}" if line else line for line in lines]

    def _list_prefix(self, oe: ET.Element) -> str:
        list_node = _child(oe, "List")
        if list_node is None:
            return ""
        number = _child(list_node, "Number")
        if number is not None:
            text = number.attrib.get("text", "").strip() or "1."
            return f"{text} "
        return "- "

    def _tag_prefix(self, oe: ET.Element) -> str:
        parts = []
        for tag in _children(oe, "Tag"):
            name = self.tag_defs.get(tag.attrib.get("index", ""), "")
            completed = tag.attrib.get("completed") == "true"
            lowered = name.casefold()
            if "to do" in lowered or "待办" in name or not name:
                box = "[x] " if completed else "[ ] "
                parts.append(box if _child(oe, "List") is not None or not self.markdown else "- " + box)
            else:
                parts.append(f"[{name}] ")
        return "".join(parts)

    def _table(self, table: ET.Element) -> list[str]:
        rows = []
        for row in _children(table, "Row"):
            cells = []
            for cell in _children(row, "Cell"):
                lines = [line.strip() for line in self._oe_children(_child(cell, "OEChildren"), "")]
                lines = [line for line in lines if line]
                if self.markdown:
                    text = "<br>".join(line.rstrip("\\") for line in lines).replace("|", "\\|")
                else:
                    text = " / ".join(lines)
                cells.append(text)
            rows.append(cells)
        if not rows:
            return []
        width = max(len(row) for row in rows)
        rows = [row + [""] * (width - len(row)) for row in rows]
        if not self.markdown:
            return [" | ".join(row) for row in rows]
        # Markdown header cells are bold already; drop the bold the writer gave them.
        header = [cell[2:-2] if re.fullmatch(r"\*\*[^*].*\*\*", cell) and "**" not in cell[2:-2] else cell for cell in rows[0]]
        rules = {"center": " :-: |", "right": " --: |"}
        lines = ["| " + " | ".join(header) + " |", "|" + "".join(rules.get(align, " --- |") for align in self._alignments(table, width))]
        lines.extend("| " + " | ".join(row) + " |" for row in rows[1:])
        return lines

    @staticmethod
    def _alignments(table: ET.Element, width: int) -> list[str]:
        """Each column's alignment, taken from its cell in the first row."""

        first = _child(table, "Row")
        cells = _children(first, "Cell") if first is not None else []
        aligns = []
        for cell in cells:
            paragraph = cell.find(f"{_NS}OEChildren/{_NS}OE")
            aligns.append(paragraph.attrib.get("alignment", "left") if paragraph is not None else "left")
        return (aligns + ["left"] * width)[:width]

    def _image(self, image: ET.Element) -> str:
        alt = " ".join(image.attrib.get("alt", "").split())
        ocr = image.find(f"{_NS}OCRData/{_NS}OCRText")
        text = (ocr.text or "").strip() if ocr is not None and self.include_ocr else ""
        if alt and text and alt.split() == text.split():
            alt = ""
        label = f"[Image: {alt}]" if alt else "[Image]"
        if not text:
            return label
        lines = [line.strip() for line in text.splitlines() if line.strip()]
        if self.markdown:
            return label + " OCR text:\n" + "\n".join(f"> {_escape(line)}\\" for line in lines[:-1]) + (
                ("\n" if len(lines) > 1 else "") + f"> {_escape(lines[-1])}"
            )
        return label + " OCR text:\n" + "\n".join(lines)

    @staticmethod
    def _file(node: ET.Element, kind: str) -> str:
        name = node.attrib.get("preferredName") or node.attrib.get("pathSource") or kind
        return f"[Attached file: {name}]"


def render_page(xml: str, fmt: ReadFormat = "text", *, include_ocr: bool = True, include_title: bool = True) -> str:
    root = ET.fromstring(xml.encode("utf-8"))
    text = PageRenderer(root, fmt, include_ocr).render(include_title=include_title)
    text = re.sub(r"[ \t]+\n", "\n", text)
    return re.sub(r"\n{3,}", "\n\n", text).strip()


def page_title(xml: str) -> str | None:
    return PageRenderer(ET.fromstring(xml.encode("utf-8")), "text", False).title()
