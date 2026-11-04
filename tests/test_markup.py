import re
import sys
import xml.etree.ElementTree as ET

import pytest

from local_onenote_mcp.image_utils import image_dimensions, proportional_dimensions
from local_onenote_mcp.markup import html_to_oe_xml, markdown_to_html
from local_onenote_mcp.pagexml import (
    build_file_page_update_xml,
    build_image_page_update_xml,
    build_page_update_xml,
    collect_page_objects,
    content_bottom,
)
from local_onenote_mcp.reader import render_page

NS = 'xmlns:one="http://schemas.microsoft.com/office/onenote/2013/onenote"'


def _page(body: str) -> str:
    return f'<one:Page {NS} ID="p"><one:Outline objectID="outline-id"><one:OEChildren>{body}</one:OEChildren></one:Outline></one:Page>'


def _texts(xml: str) -> list[str]:
    return re.findall(r"<!\[CDATA\[(.*?)\]\]>", xml, re.S)


def test_page_update_is_well_formed_and_uses_cdata():
    xml = build_page_update_xml("page-id", title="A & B", content="Hello\nWorld")

    ET.fromstring(xml)
    assert "<![CDATA[A &amp; B]]>" in xml
    assert _texts(xml)[1:] == ["Hello", "World"]


def test_markdown_whitespace_between_blocks_does_not_create_blank_paragraphs():
    xml = build_page_update_xml("p", content="# Title\n\nIntro\n\n- a\n- b\n", content_format="markdown")

    assert all(text.strip() for text in _texts(xml))


def test_markdown_lists_become_native_nested_onenote_lists():
    xml = html_to_oe_xml(markdown_to_html("- one\n- two\n  - nested\n\n1. first\n2. second\n"))

    assert xml.count('<one:Bullet bullet="2"/>') == 2
    assert xml.count('<one:Bullet bullet="3"/>') == 1
    assert '<one:Number numberSequence="0" numberFormat="##." text="2."/>' in xml
    assert re.search(r"two\]\]></one:T><one:OEChildren><one:OE><one:List><one:Bullet", xml)


def test_markdown_task_items_and_code_blocks_without_native_styles():
    xml = html_to_oe_xml(markdown_to_html("- [ ] todo\n- [x] done\n\n```\ndef f():\n    return 1\n```\n"))
    texts = _texts(xml)

    assert "☐ todo" in texts and "☑ done" in texts
    assert any("&#160;&#160;&#160;&#160;return 1" in text for text in texts)


def test_page_update_uses_native_headings_todo_tags_and_code_style():
    xml = build_page_update_xml("p", content="# Plan\n\n- [x] done\n- [ ] open\n\n```\nx = 1\n```\n", content_format="markdown")

    assert '<one:TagDef index="0" type="0" symbol="3"' in xml
    assert '<one:QuickStyleDef index="0" name="h1"' in xml
    assert '<one:QuickStyleDef index="1" name="code"' in xml
    assert '<one:OE quickStyleIndex="0"><one:T><![CDATA[Plan]]>' in xml
    assert '<one:Tag index="0" completed="true"/><one:T><![CDATA[done]]>' in xml
    assert xml.index("<one:TagDef") < xml.index("<one:QuickStyleDef") < xml.index("<one:Outline")


def test_inline_fallback_without_native_styles():
    xml = build_page_update_xml("p", content="# Heading", content_format="markdown", native_styles=False)

    assert '<span style="font-size:20.0pt;font-weight:bold">Heading</span>' in xml
    assert "QuickStyleDef" not in xml


def test_plain_page_update_adds_no_style_definitions():
    xml = build_page_update_xml("p", title="T", content="hello")

    assert "QuickStyleDef" not in xml and "TagDef" not in xml


def test_html_table_becomes_native_table_padded_to_column_count():
    xml = build_page_update_xml(
        "p",
        content="<table><tr><th>A</th><th>B</th><th>C</th></tr><tr><td><strong>1</strong></td><td>2</td></tr></table>",
        content_format="html",
    )

    assert '<one:Table bordersVisible="true"' in xml
    assert xml.count("<one:Column index=") == 3 and xml.count("<one:Cell") == 6
    assert 'shadingColor="#D9EAF7"' in xml and "<![CDATA[<strong>1</strong>]]>" in xml


def test_table_columns_follow_their_text_and_fit_the_page():
    import re

    def widths(markdown: str) -> list[float]:
        xml = build_page_update_xml("p", content=markdown, content_format="markdown")
        return [float(w) for w in re.findall(r'<one:Column index="\d+" width="([\d.]+)"', xml)]

    small = widths("| a | b |\n|---|---|\n| 1 | NUS Multipurpose Sports Hall (MPSH 2) |")
    assert small[0] < 60 and small[1] > 3 * small[0]
    assert widths("| 考场 |\n|---|\n| 体育馆 |")[0] < widths("| 考场 |\n|---|\n| 多功能体育馆二号馆 |")[0]
    wide = widths("|" + " long column text here |" * 8 + "\n" + "|---" * 8 + "|")
    assert sum(wide) <= 960.5


def test_html_writer_drops_scripts_and_unsafe_links():
    xml = html_to_oe_xml('<p><a href="javascript:alert(1)">x</a><script>bad()</script> ok <s>gone</s></p>')

    assert "javascript" not in xml and "bad()" not in xml
    assert "ok" in xml and "text-decoration:line-through" in xml


def test_html_paragraph_style_reaches_each_line():
    texts = _texts(html_to_oe_xml('<p style="color:red">one<br>two</p>'))

    assert texts == ['<span style="color:red">one</span>', '<span style="color:red">two</span>']

def test_image_and_file_xml():
    image = build_image_page_update_xml("p", image_base64="abc", image_format="PNG", width=320, height=180)
    attached = build_file_page_update_xml("p", path='C:/a "b".pdf', name="b.pdf", x=36, y=200)

    assert '<one:Image format="png"><one:Size width="320.00" height="180.00"/>' in image
    assert "<one:Size" not in build_image_page_update_xml("p", image_base64="abc", image_format="png")
    assert 'pathSource="C:/a &quot;b&quot;.pdf"' in attached and '<one:Position x="36.00" y="200.00"' in attached


def test_png_dimensions_support_proportional_image_sizing(tmp_path):
    image_path = tmp_path / "sample.png"
    image_path.write_bytes(b"\x89PNG\r\n\x1a\n\x00\x00\x00\rIHDR\x00\x00\x02\x00\x00\x00\x01\x00\x08\x02\x00\x00\x00\x00\x00\x00\x00")

    assert image_dimensions(image_path) == (512, 256)
    assert proportional_dimensions(image_path, width=256, height=None) == (256.0, 128.0)


def test_new_blocks_go_below_existing_content():
    xml = (
        f'<one:Page {NS} ID="p">'
        '<one:Outline><one:Position x="36" y="80"/><one:Size width="300" height="120"/></one:Outline>'
        '<one:Image><one:Position x="36" y="210"/><one:Size width="100" height="50"/></one:Image>'
        "</one:Page>"
    )

    assert content_bottom(xml) == 260.0
    assert content_bottom(f'<one:Page {NS} ID="p"><one:Title/></one:Page>') is None


def test_collect_page_objects_reports_callbacks_and_deletable_blocks():
    xml = _page("<one:OE objectID='oe-id'><one:Image format='png'><one:CallbackID callbackID='cb-1'/></one:Image></one:OE>")
    xml = xml.replace('objectID="outline-id">', 'objectID="outline-id"><one:Position x="36" y="80"/><one:Size width="300" height="120"/>')
    objects = {obj["type"]: obj for obj in collect_page_objects(xml)}

    assert set(objects) == {"Outline", "Image"}
    assert objects["Outline"]["delete_supported"] is True
    assert (objects["Outline"]["x"], objects["Outline"]["y"], objects["Outline"]["height"]) == (36.0, 80.0, 120.0)
    assert objects["Image"]["callback_id"] == "cb-1" and objects["Image"]["delete_object_id"] == "outline-id"
    assert objects["Image"]["delete_supported"] is False and objects["Image"]["parent_object_id"] == "outline-id"


def test_reader_maps_native_styles_back_to_markdown():
    xml = (
        f'<one:Page {NS} ID="p">'
        '<one:TagDef index="0" type="0" symbol="3" name="To Do"/>'
        '<one:QuickStyleDef index="1" name="h2"/><one:QuickStyleDef index="2" name="code"/>'
        "<one:Outline><one:OEChildren>"
        '<one:OE quickStyleIndex="1"><one:T><![CDATA[Section]]></one:T></one:OE>'
        '<one:OE><one:Tag index="0" completed="true"/><one:T><![CDATA[shipped]]></one:T></one:OE>'
        '<one:OE quickStyleIndex="2"><one:T><![CDATA[a = 1]]></one:T></one:OE>'
        '<one:OE quickStyleIndex="2"><one:T><![CDATA[b = 2]]></one:T></one:OE>'
        "</one:OEChildren></one:Outline></one:Page>"
    )

    assert render_page(xml, "markdown") == "## Section\n- [x] shipped\n```\na = 1\nb = 2\n```"


def test_reader_renders_lists_tables_and_image_ocr_as_markdown():
    xml = _page(
        "<one:OE><one:List><one:Bullet bullet='2'/></one:List>"
        "<one:T><![CDATA[<span style='font-weight:bold'>Item</span>]]></one:T>"
        "<one:OEChildren><one:OE><one:List><one:Number text='1.'/></one:List>"
        "<one:T><![CDATA[child]]></one:T></one:OE></one:OEChildren></one:OE>"
        "<one:OE><one:Table>"
        "<one:Row><one:Cell><one:OEChildren><one:OE><one:T><![CDATA[H]]></one:T></one:OE></one:OEChildren></one:Cell></one:Row>"
        "<one:Row><one:Cell><one:OEChildren><one:OE><one:T><![CDATA[a|b]]></one:T></one:OE></one:OEChildren></one:Cell></one:Row>"
        "</one:Table></one:OE>"
        "<one:OE><one:Image><one:OCRData><one:OCRText><![CDATA[scanned words]]></one:OCRText></one:OCRData></one:Image></one:OE>"
    )

    text = render_page(xml, "markdown")

    assert "- **Item**\n  1. child" in text
    assert "| H |\n| --- |\n| a\\|b |" in text
    assert "> scanned words" in text
    assert "scanned words" not in render_page(xml, "text", include_ocr=False)


def test_reader_merges_font_runs_before_adding_marks():
    yahei, latin = "font-family:\"Microsoft YaHei\"", "font-family:Calibri"
    xml = (
        f'<one:Page {NS} ID="p"><one:QuickStyleDef index="0" name="quote" italic="true"/>'
        "<one:Outline><one:OEChildren>"
        f"<one:OE quickStyleIndex='0'><one:T><![CDATA[<span style='{yahei};font-style:italic'>第三方</span>"
        f"<span style='{latin};font-style:italic'> API </span><span style='{yahei};font-style:italic'>限流</span>]]></one:T></one:OE>"
        f"<one:OE><one:T><![CDATA[<span style='{yahei};text-decoration:line-through'>删除线</span>"
        f"<span style='{latin};text-decoration:line-through'> </span><span style='{yahei}'>和 </span>"
        f"<span style='{yahei};font-weight:bold'>粗</span><span style='{latin};font-weight:bold'>体</span>]]></one:T></one:OE>"
        "<one:OE><one:List><one:Number text='1.'/></one:List><one:T><![CDATA[step]]></one:T>"
        "<one:OEChildren><one:OE><one:List><one:Bullet bullet='2'/></one:List><one:T><![CDATA[detail]]></one:T></one:OE></one:OEChildren></one:OE>"
        "</one:OEChildren></one:Outline></one:Page>"
    )

    assert render_page(xml, "markdown").split("\n") == ["> 第三方 API 限流", "~~删除线~~ 和 **粗体**", "1. step", "   - detail"]


def test_reader_uses_recognised_ink_text_and_labels_the_rest():
    xml = _page(
        "<one:OE><one:InkWord recognizedText='hello'/></one:OE><one:OE><one:InkWord/></one:OE>"
    ).replace("</one:Page>", "<one:InkDrawing objectID='ink'/></one:Page>")

    text = render_page(xml, "text")

    assert text.startswith("hello\n[Handwriting:")
    assert text.count("[Handwriting:") == 2


@pytest.mark.skipif(sys.platform != "win32", reason="PowerShell bridge is Windows-only")
def test_bridge_round_trips_unicode_and_restarts():
    from local_onenote_mcp.bridge import OneNoteBridge

    bridge = OneNoteBridge(timeout_seconds=30)
    try:
        value = '中文 "quoted"\nline ✓'
        assert bridge.call("ping", value=value)["value"] == value
        bridge.close()
        assert bridge.call("ping", value="again")["value"] == "again"
    finally:
        bridge.close()
