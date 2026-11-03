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


class _InlineReader(HTMLParser):
    """Turn a one:T inline HTML fragment into text, optionally with Markdown marks.

    OneNote splits text into runs whenever the font changes (Chinese and Latin text get separate runs), so the
    Markdown output merges neighbouring runs with the same formatting before adding marks.
    """

    def __init__(self, markdown: bool, ignore: str = "") -> None:
        super().__init__(convert_charrefs=True)
        self.markdown = markdown
        self.ignore = ignore
        self.segments: list[tuple[str, str, str]] = []  # (text, marks, link)
        self._marks: list[set[str]] = []
        self._links: list[str] = []

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        tag = tag.lower()
        if tag == "br":
            self.handle_data("\n")
            return
        attributes = {name.lower(): value or "" for name, value in attrs}
        if tag == "a":
            self._links.append(attributes.get("href", ""))
            return
        style = attributes.get("style", "").replace(" ", "").lower()
        marks = set()
        if tag in {"b", "strong"} or "font-weight:bold" in style:
            marks.add("**")
        if tag in {"i", "em"} or "font-style:italic" in style:
            marks.add("*")
        if tag in {"s", "del", "strike"} or "line-through" in style:
            marks.add("~~")
        if "consolas" in style or "courier" in style:
            marks.add("`")
        self._marks.append(marks)

    def handle_endtag(self, tag: str) -> None:
        tag = tag.lower()
        if tag == "a":
            if self._links:
                self._links.pop()
        elif tag != "br" and self._marks:
            self._marks.pop()

    def handle_data(self, data: str) -> None:
        held = set().union(*self._marks)
        marks = "".join(mark for mark in ("**", "*", "~~", "`") if mark in held and mark not in self.ignore)
        link = self._links[-1] if self._links else ""
        data = data.replace("\xa0", " ")
        if self.segments and (self.segments[-1][1:] == (marks, link) or (not data.strip() and not link)):
            previous, marks, link = self.segments.pop()
            data = previous + data
        self.segments.append((data, marks, link))

    def result(self) -> str:
        if not self.markdown:
            return "".join(text for text, _, _ in self.segments).replace("\x00", "")
        out = []
        for text, marks, link in self.segments:
            core = text.strip()
            if core and marks:
                lead, trail = text[: len(text) - len(text.lstrip())], text[len(text.rstrip()) :]
                text = f"{lead}{marks}{core}{marks[::-1]}{trail}"
            if link and core:
                text = f"[{text}]({link})"
            out.append(text)
        return "".join(out).replace("\x00", "")


def inline_text(fragment: str | None, *, markdown: bool = False, ignore: str = "") -> str:
    parser = _InlineReader(markdown, ignore)
    parser.feed(fragment or "")
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
        if container is None:
            return []
        lines: list[str] = []
        in_code = False
        for oe in _children(container, "OE"):
            is_code = self.markdown and self._style(oe) == "code"
            if is_code != in_code:
                lines.append(indent + "```")
                in_code = is_code
            lines.extend(self._oe(oe, indent))
        if in_code:
            lines.append(indent + "```")
        return lines

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
        child_indent = indent + (" " * len(marker) if marker and self.markdown else "  ")
        code = self.markdown and self._style(oe) == "code"
        quote = style_prefix == "> "
        lines: list[str] = []
        body: list[str] = []
        for node in oe:
            kind = _local(node.tag)
            if kind == "T":
                body.append(inline_text(node.text, markdown=self.markdown and not code, ignore="*" if quote else ""))
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
            first, *rest = text.split("\n")
            text_lines = [f"{indent}{prefix}{first}".rstrip()]
            continuation = indent + " " * len(prefix)
            text_lines.extend(f"{continuation}{line}".rstrip() for line in rest)
            lines = text_lines + lines
        elif not lines and not self._has_children(oe):
            lines.append("")
        lines.extend(self._oe_children(_child(oe, "OEChildren"), child_indent))
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
                text = "\n".join(line.strip() for line in self._oe_children(_child(cell, "OEChildren"), ""))
                text = text.strip()
                if self.markdown:
                    text = text.replace("|", "\\|").replace("\n", "<br>")
                else:
                    text = text.replace("\n", " / ")
                cells.append(text)
            rows.append(cells)
        if not rows:
            return []
        width = max(len(row) for row in rows)
        rows = [row + [""] * (width - len(row)) for row in rows]
        if not self.markdown:
            return [" | ".join(row) for row in rows]
        lines = ["| " + " | ".join(rows[0]) + " |", "|" + " --- |" * width]
        lines.extend("| " + " | ".join(row) + " |" for row in rows[1:])
        return lines

    def _image(self, image: ET.Element) -> str:
        alt = image.attrib.get("alt", "").strip()
        label = f"[Image: {alt}]" if alt else "[Image]"
        if not self.include_ocr:
            return label
        ocr = image.find(f"{_NS}OCRData/{_NS}OCRText")
        text = (ocr.text or "").strip() if ocr is not None else ""
        if not text:
            return label
        lines = [line.strip() for line in text.splitlines() if line.strip()]
        if self.markdown:
            return label + " OCR text:\n" + "\n".join(f"> {line}" for line in lines)
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
