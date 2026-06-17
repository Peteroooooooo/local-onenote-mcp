"""OneNote XML and inline HTML helpers."""

from __future__ import annotations

import html
import os
import re
import subprocess
import tempfile
import winreg
import xml.etree.ElementTree as ET
from dataclasses import dataclass
from html.parser import HTMLParser
from pathlib import Path
from typing import Any, Iterable, Literal

from .constants import ONE_NS

ET.register_namespace("one", ONE_NS)


BLOCK_TAGS = {
    "address",
    "article",
    "aside",
    "blockquote",
    "div",
    "figcaption",
    "figure",
    "footer",
    "h1",
    "h2",
    "h3",
    "h4",
    "h5",
    "h6",
    "header",
    "li",
    "main",
    "nav",
    "p",
    "pre",
    "section",
    "td",
    "th",
    "tr",
}

INLINE_TAGS = {
    "a",
    "b",
    "br",
    "code",
    "del",
    "em",
    "i",
    "mark",
    "span",
    "strike",
    "strong",
    "sub",
    "sup",
    "s",
    "u",
}

SAFE_ATTRS = {
    "a": {"href", "title"},
    "span": {"style"},
}

INLINE_STYLE_TAGS = {
    "code": "font-family:Consolas,'Courier New',monospace",
    "del": "text-decoration:line-through",
    "mark": "background:#FFF2CC",
    "s": "text-decoration:line-through",
    "strike": "text-decoration:line-through",
}

DELETABLE_PAGE_OBJECT_TYPES = {"Outline", "Image", "InkDrawing", "FileAttachment", "InsertedFile", "MediaFile"}

HEADING_STYLES = {
    "h1": "font-size:20.0pt;font-weight:bold",
    "h2": "font-size:16.0pt;font-weight:bold",
    "h3": "font-size:14.0pt;font-weight:bold",
    "h4": "font-size:12.0pt;font-weight:bold",
    "h5": "font-size:11.0pt;font-weight:bold",
    "h6": "font-size:10.0pt;font-weight:bold",
}


class InlineHTMLSanitizer(HTMLParser):
    """Convert arbitrary HTML-ish input to a OneNote-friendly inline fragment."""

    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.parts: list[str] = []
        self._drop_stack: list[str] = []

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        tag = tag.lower()
        if tag in {"script", "style"}:
            self._drop_stack.append(tag)
            return
        if self._drop_stack:
            return
        if tag in HEADING_STYLES:
            self._append_break()
            self.parts.append(f'<span style="{HEADING_STYLES[tag]}">')
            return
        if tag in BLOCK_TAGS:
            self._append_break()
            return
        if tag not in INLINE_TAGS:
            return
        if tag == "br":
            self._append_break()
            return
        if tag in INLINE_STYLE_TAGS:
            self.parts.append(f'<span style="{INLINE_STYLE_TAGS[tag]}">')
            return
        allowed = SAFE_ATTRS.get(tag, set())
        rendered_attrs = []
        for name, value in attrs:
            if value is None:
                continue
            name = name.lower()
            if name not in allowed:
                continue
            if name == "href" and not value.lower().startswith(("http://", "https://", "onenote:", "mailto:")):
                continue
            rendered_attrs.append(f'{name}="{html.escape(value, quote=True)}"')
        attr_text = (" " + " ".join(rendered_attrs)) if rendered_attrs else ""
        self.parts.append(f"<{tag}{attr_text}>")

    def handle_endtag(self, tag: str) -> None:
        tag = tag.lower()
        if self._drop_stack:
            if tag == self._drop_stack[-1]:
                self._drop_stack.pop()
            return
        if tag in HEADING_STYLES:
            self.parts.append("</span>")
            self._append_break()
            return
        if tag in BLOCK_TAGS:
            self._append_break()
            return
        if tag in INLINE_STYLE_TAGS:
            self.parts.append("</span>")
            return
        if tag in INLINE_TAGS and tag != "br":
            self.parts.append(f"</{tag}>")

    def handle_data(self, data: str) -> None:
        if not self._drop_stack:
            self.parts.append(html.escape(data, quote=False))

    def get_html(self) -> str:
        text = "".join(self.parts)
        text = re.sub(r"(?:<br/>){3,}", "<br/><br/>", text)
        text = re.sub(r"^(?:<br/>)+|(?:<br/>)+$", "", text)
        return text.strip()

    def _append_break(self) -> None:
        if not self.parts or self.parts[-1] != "<br/>":
            self.parts.append("<br/>")


@dataclass
class TextBlock:
    html: str


@dataclass
class TableCell:
    html: str
    header: bool = False


@dataclass
class TableBlock:
    rows: list[list[TableCell]]


ContentBlock = TextBlock | TableBlock


class OneNoteHTMLBlockParser(HTMLParser):
    """Convert simple HTML into ordered text/table blocks for OneNote XML.

    OneNote's COM API does not preserve HTML <table> markup when it is put
    inside one:T CDATA. Tables must be emitted as native one:Table elements.
    """

    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.blocks: list[ContentBlock] = []
        self._text = InlineHTMLSanitizer()
        self._table_depth = 0
        self._rows: list[list[TableCell]] = []
        self._current_row: list[TableCell] | None = None
        self._current_cell: InlineHTMLSanitizer | None = None
        self._current_cell_header = False
        self._drop_stack: list[str] = []

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        tag = tag.lower()
        if tag == "table":
            if self._table_depth == 0:
                self._flush_text()
                self._rows = []
                self._current_row = None
                self._current_cell = None
            self._table_depth += 1
            return
        if self._table_depth:
            self._handle_table_starttag(tag, attrs)
            return
        self._text.handle_starttag(tag, attrs)

    def handle_endtag(self, tag: str) -> None:
        tag = tag.lower()
        if tag == "table" and self._table_depth:
            self._close_cell()
            self._close_row()
            self._table_depth -= 1
            if self._table_depth == 0:
                rows = [row for row in self._rows if row]
                if rows:
                    self.blocks.append(TableBlock(rows=rows))
            return
        if self._table_depth:
            self._handle_table_endtag(tag)
            return
        self._text.handle_endtag(tag)

    def handle_data(self, data: str) -> None:
        if self._table_depth:
            if not self._drop_stack and self._current_cell is not None:
                self._current_cell.handle_data(data)
            return
        self._text.handle_data(data)

    def get_blocks(self) -> list[ContentBlock]:
        self._flush_text()
        return self.blocks

    def _handle_table_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        if tag in {"script", "style"}:
            self._drop_stack.append(tag)
            return
        if self._drop_stack:
            return
        if tag == "tr":
            self._close_cell()
            self._close_row()
            self._current_row = []
            return
        if tag in {"td", "th"}:
            if self._current_row is None:
                self._current_row = []
            self._close_cell()
            self._current_cell = InlineHTMLSanitizer()
            self._current_cell_header = tag == "th"
            return
        if self._current_cell is not None:
            self._current_cell.handle_starttag(tag, attrs)
            return

    def _handle_table_endtag(self, tag: str) -> None:
        if self._drop_stack:
            if tag == self._drop_stack[-1]:
                self._drop_stack.pop()
            return
        if tag in {"td", "th"}:
            self._close_cell()
            return
        if tag == "tr":
            self._close_cell()
            self._close_row()
            return
        if self._current_cell is not None:
            self._current_cell.handle_endtag(tag)

    def _flush_text(self) -> None:
        text = self._text.get_html()
        if text:
            self.blocks.append(TextBlock(html=text))
        self._text = InlineHTMLSanitizer()

    def _close_row(self) -> None:
        if self._current_row:
            self._rows.append(self._current_row)
        self._current_row = None

    def _close_cell(self) -> None:
        if self._current_cell is None:
            return
        cell_html = self._cell_html()
        if cell_html or self._current_row is not None:
            if self._current_row is None:
                self._current_row = []
            self._current_row.append(TableCell(html=cell_html, header=self._current_cell_header))
        self._current_cell = None
        self._current_cell_header = False

    def _cell_html(self) -> str:
        if self._current_cell is None:
            return ""
        return self._current_cell.get_html()

    def _append_cell_break(self) -> None:
        if self._current_cell is None:
            return
        self._current_cell.handle_starttag("br", [])


class HTMLTextExtractor(HTMLParser):
    """Extract readable text from a OneNote T element's inline HTML fragment."""

    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.parts: list[str] = []

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        if tag.lower() in {"br", "p", "div", "li", "tr", "h1", "h2", "h3", "h4", "h5", "h6"}:
            self._newline()

    def handle_endtag(self, tag: str) -> None:
        if tag.lower() in {"p", "div", "li", "tr", "h1", "h2", "h3", "h4", "h5", "h6"}:
            self._newline()

    def handle_data(self, data: str) -> None:
        if data:
            self.parts.append(data)

    def text(self) -> str:
        value = "".join(self.parts)
        value = value.replace("\x00", "")
        value = re.sub(r"[ \t]+\n", "\n", value)
        value = re.sub(r"\n{3,}", "\n\n", value)
        return value.strip()

    def _newline(self) -> None:
        if not self.parts or not self.parts[-1].endswith("\n"):
            self.parts.append("\n")


def normalize_content(content: str, content_format: str = "plain") -> str:
    """Return OneNote inline HTML for plain text or simple HTML input."""

    if content_format == "plain":
        escaped = html.escape(content, quote=False)
        return escaped.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "<br/>")
    if content_format == "html":
        parser = InlineHTMLSanitizer()
        parser.feed(content)
        parser.close()
        return parser.get_html()
    if content_format in {"markdown", "md"}:
        return normalize_content(markdown_to_html(content), "html")
    raise ValueError("content_format must be 'plain', 'html', or 'markdown'.")


def _registry_value(root: int, subkey: str, value_name: str) -> str | None:
    try:
        with winreg.OpenKey(root, subkey) as key:
            value, _ = winreg.QueryValueEx(key, value_name)
            return str(value) if value else None
    except OSError:
        return None


def find_markdig_dll() -> Path:
    """Locate OneMore's bundled Markdig Markdown parser."""

    env_path = os.environ.get("LOCAL_ONENOTE_MARKDIG_DLL")
    candidates: list[Path] = []
    if env_path:
        candidates.append(Path(env_path))

    addin_path = _registry_value(
        winreg.HKEY_CURRENT_USER,
        r"Software\Microsoft\Windows\CurrentVersion\App Paths\River.OneMoreAddIn.dll",
        "Path",
    )
    if addin_path:
        candidates.append(Path(addin_path).parent / "Markdig.Signed.dll")

    candidates.extend(
        [
            Path(r"C:\Program Files\River\OneMoreAddIn\Markdig.Signed.dll"),
            Path(r"C:\Program Files (x86)\River\OneMoreAddIn\Markdig.Signed.dll"),
        ]
    )
    for candidate in candidates:
        if candidate.is_file():
            return candidate
    raise RuntimeError(
        "Markdown conversion requires OneMore's Markdig.Signed.dll. "
        "Install OneMore or set LOCAL_ONENOTE_MARKDIG_DLL."
    )


MARKDOWN_POWERSHELL = r'''
$ErrorActionPreference = "Stop"
if ([string]::IsNullOrWhiteSpace($env:LOCAL_ONENOTE_MARKDIG_DLL)) {
    throw "LOCAL_ONENOTE_MARKDIG_DLL is not set."
}
if ([string]::IsNullOrWhiteSpace($env:LOCAL_ONENOTE_MARKDOWN_INPUT) -or [string]::IsNullOrWhiteSpace($env:LOCAL_ONENOTE_MARKDOWN_OUTPUT)) {
    throw "Markdown input/output paths are not set."
}
[Reflection.Assembly]::LoadFrom($env:LOCAL_ONENOTE_MARKDIG_DLL) | Out-Null
$builder = [Markdig.MarkdownPipelineBuilder]::new()
[Markdig.MarkdownExtensions]::UseAdvancedExtensions($builder) | Out-Null
$pipeline = $builder.Build()
$markdown = Get-Content -LiteralPath $env:LOCAL_ONENOTE_MARKDOWN_INPUT -Raw -Encoding UTF8
$writer = [System.IO.StringWriter]::new()
[Markdig.Markdown]::ToHtml($markdown, $writer, $pipeline, $null) | Out-Null
$html = $writer.ToString()
[System.IO.File]::WriteAllText($env:LOCAL_ONENOTE_MARKDOWN_OUTPUT, $html, [System.Text.UTF8Encoding]::new($false))
if (!(Test-Path -LiteralPath $env:LOCAL_ONENOTE_MARKDOWN_OUTPUT)) {
    throw "Markdown conversion did not write an HTML output file."
}
Write-Host "markdown-converted"
'''


def markdown_to_html(content: str) -> str:
    """Convert Markdown to HTML through OneMore's bundled Markdig parser."""

    markdig_dll = find_markdig_dll()
    script_path = _write_temp_text(MARKDOWN_POWERSHELL, ".ps1")
    input_path = _write_temp_text(content, ".md")
    output_path = _reserve_temp_path(".html")
    env = os.environ.copy()
    env["LOCAL_ONENOTE_MARKDIG_DLL"] = str(markdig_dll)
    env["LOCAL_ONENOTE_MARKDOWN_INPUT"] = str(input_path)
    env["LOCAL_ONENOTE_MARKDOWN_OUTPUT"] = str(output_path)
    try:
        completed = subprocess.run(
            ["powershell.exe", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-File", str(script_path)],
            text=True,
            capture_output=True,
            timeout=int(os.environ.get("LOCAL_ONENOTE_MARKDOWN_TIMEOUT", "30")),
            env=env,
        )
        if completed.returncode != 0:
            message = completed.stderr.strip() or completed.stdout.strip() or "Markdown conversion failed."
            raise RuntimeError(message)
        if not output_path.exists():
            raise RuntimeError("Markdown conversion did not write an HTML output file.")
        return output_path.read_text(encoding="utf-8-sig")
    finally:
        _remove_quietly(script_path)
        _remove_quietly(input_path)
        _remove_quietly(output_path)


def _write_temp_text(value: str, suffix: str) -> Path:
    handle = tempfile.NamedTemporaryFile(
        mode="w",
        encoding="utf-8",
        prefix="local-onenote-mcp-",
        suffix=suffix,
        delete=False,
    )
    with handle:
        handle.write(value)
    return Path(handle.name)


def _reserve_temp_path(suffix: str) -> Path:
    handle = tempfile.NamedTemporaryFile(prefix="local-onenote-mcp-", suffix=suffix, delete=False)
    path = Path(handle.name)
    handle.close()
    path.unlink(missing_ok=True)
    return path


def _remove_quietly(path: Path) -> None:
    try:
        path.unlink(missing_ok=True)
    except OSError:
        pass


def html_fragment_to_text(fragment: str) -> str:
    parser = HTMLTextExtractor()
    parser.feed(fragment or "")
    parser.close()
    return parser.text()


def cdata(value: str) -> str:
    return "<![CDATA[" + value.replace("]]>", "]]]]><![CDATA[>") + "]]>"


def attr(value: str) -> str:
    return html.escape(value, quote=True)


def one_t(fragment_html: str) -> str:
    return f"<one:T>{cdata(fragment_html)}</one:T>"


def oe_children(fragment_html: str) -> str:
    parts = re.split(r"<br\s*/?>", fragment_html)
    if not parts:
        parts = [""]
    return "".join(f"<one:OE>{one_t(part)}</one:OE>" for part in parts)


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
        f'<one:OE alignment="left" quickStyleIndex="0"{style_attr}>{one_t(cell_html)}</one:OE>'
        "</one:OEChildren>"
        "</one:Cell>"
    )


def _table_column_widths(rows: list[list[TableCell]]) -> list[float]:
    column_count = max((len(row) for row in rows), default=0)
    if column_count <= 0:
        return []
    # Keep tables comfortably inside a normal OneNote page width while still
    # allowing enough room for short checklist-style tables.
    total_width = 960.0
    width = max(90.0, min(220.0, total_width / column_count))
    return [width] * column_count


def one_table(rows: list[list[TableCell]]) -> str:
    widths = _table_column_widths(rows)
    if not widths:
        return ""
    column_xml = "".join(
        f'<one:Column index="{index}" width="{width:.1f}" isLocked="true"/>'
        for index, width in enumerate(widths)
    )
    row_xml = []
    column_count = len(widths)
    for row in rows:
        padded = row + [TableCell(html="")] * (column_count - len(row))
        row_xml.append("<one:Row>" + "".join(_one_table_cell(cell) for cell in padded[:column_count]) + "</one:Row>")
    return (
        '<one:OE alignment="left"><one:Table bordersVisible="true" hasHeaderRow="false">'
        f"<one:Columns>{column_xml}</one:Columns>"
        f"{''.join(row_xml)}"
        "</one:Table></one:OE>"
    )


def html_content_blocks(content: str) -> list[ContentBlock]:
    parser = OneNoteHTMLBlockParser()
    parser.feed(content)
    parser.close()
    return parser.get_blocks()


def content_to_oe_xml(content: str, content_format: Literal["plain", "html", "markdown", "md"] = "plain") -> str:
    if content_format == "plain":
        return oe_children(normalize_content(content, "plain"))
    if content_format == "html":
        blocks = html_content_blocks(content)
        if not blocks:
            return ""
        parts: list[str] = []
        for block in blocks:
            if isinstance(block, TextBlock):
                parts.append(oe_children(block.html))
            else:
                parts.append(one_table(block.rows))
        return "".join(parts)
    if content_format in {"markdown", "md"}:
        return content_to_oe_xml(markdown_to_html(content), "html")
    raise ValueError("content_format must be 'plain', 'html', or 'markdown'.")


def build_outline_xml(
    content: str,
    *,
    content_format: Literal["plain", "html", "markdown", "md"] = "plain",
    object_id: str | None = None,
    x: float | None = None,
    y: float | None = None,
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
        f"<one:OEChildren>{content_to_oe_xml(content, content_format)}</one:OEChildren>"
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
) -> str:
    parts = [f'<one:Page xmlns:one="{ONE_NS}" ID="{attr(page_id)}">']
    if title is not None:
        parts.append(build_title_xml(title))
    if content is not None and content != "":
        parts.append(build_outline_xml(content, content_format=content_format, x=x, y=y))
    parts.append("</one:Page>")
    return "".join(parts)


def build_image_page_update_xml(
    page_id: str,
    *,
    image_base64: str,
    image_format: str,
    x: float = 36.0,
    y: float = 120.0,
    width: float | None = None,
    height: float | None = None,
) -> str:
    size = ""
    if width is not None and height is not None:
        size = f'<one:Size width="{float(width):.2f}" height="{float(height):.2f}"/>'
    return (
        f'<one:Page xmlns:one="{ONE_NS}" ID="{attr(page_id)}">'
        "<one:Outline>"
        f'<one:Position x="{float(x):.2f}" y="{float(y):.2f}" z="0"/>'
        "<one:OEChildren><one:OE>"
        f'<one:Image format="{attr(image_format.lower())}">'
        f"{size}<one:Data>{image_base64}</one:Data>"
        "</one:Image>"
        "</one:OE></one:OEChildren>"
        "</one:Outline>"
        "</one:Page>"
    )


def local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1] if "}" in tag else tag


def parse_xml(xml: str) -> ET.Element:
    return ET.fromstring(xml.encode("utf-8"))


def text_from_page_xml(xml: str) -> str:
    root = parse_xml(xml)
    texts = []
    for node in root.iter():
        if local_name(node.tag) == "T" and node.text:
            texts.append(html_fragment_to_text(node.text))
    return "\n\n".join(t for t in texts if t).strip()


def title_from_page_xml(xml: str) -> str | None:
    root = parse_xml(xml)
    for title in root.iter():
        if local_name(title.tag) != "Title":
            continue
        for node in title.iter():
            if local_name(node.tag) == "T" and node.text:
                value = html_fragment_to_text(node.text)
                if value:
                    return value
    return None


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
            if "callbackID" in node.attrib:
                record["callback_id"] = node.attrib["callbackID"]
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
