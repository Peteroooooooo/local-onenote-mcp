import re
import sys

import pytest

from local_onenote_mcp.markup import html_to_oe_xml, markdown_to_html
from local_onenote_mcp.reader import render_page
from local_onenote_mcp.xml_utils import build_page_update_xml, collect_page_objects

NS = 'xmlns:one="http://schemas.microsoft.com/office/onenote/2013/onenote"'


def _page(body: str) -> str:
    return f'<one:Page {NS} ID="p"><one:Outline><one:OEChildren>{body}</one:OEChildren></one:Outline></one:Page>'


def _texts(xml: str) -> list[str]:
    return re.findall(r"<!\[CDATA\[(.*?)\]\]>", xml, re.S)


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

    assert "☐ todo" in texts
    assert "☑ done" in texts
    assert any("&#160;&#160;&#160;&#160;return 1" in text for text in texts)


def test_page_update_uses_native_headings_todo_tags_and_code_style():
    xml = build_page_update_xml("p", content="# Plan\n\n- [x] done\n- [ ] open\n\n```\nx = 1\n```\n", content_format="markdown")

    assert '<one:TagDef index="0" type="0" symbol="3"' in xml
    assert '<one:QuickStyleDef index="1" name="h1"' in xml
    assert '<one:QuickStyleDef index="2" name="code"' in xml
    assert '<one:OE quickStyleIndex="1"><one:T><![CDATA[Plan]]>' in xml
    assert '<one:Tag index="0" completed="true"/><one:T><![CDATA[done]]>' in xml
    assert '<one:Tag index="0" completed="false"/><one:T><![CDATA[open]]>' in xml
    assert xml.index("<one:TagDef") < xml.index("<one:Outline")


def test_plain_page_update_adds_no_style_definitions():
    xml = build_page_update_xml("p", title="T", content="hello")

    assert "QuickStyleDef" not in xml and "TagDef" not in xml


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


def test_new_blocks_go_below_existing_content():
    from local_onenote_mcp.xml_utils import content_bottom

    xml = (
        f'<one:Page {NS} ID="p">'
        '<one:Outline><one:Position x="36" y="80"/><one:Size width="300" height="120"/></one:Outline>'
        '<one:Image><one:Position x="36" y="210"/><one:Size width="100" height="50"/></one:Image>'
        "</one:Page>"
    )

    assert content_bottom(xml) == 260.0
    assert content_bottom(f'<one:Page {NS} ID="p"><one:Title/></one:Page>') is None


def test_html_writer_drops_scripts_and_unsafe_links():
    xml = html_to_oe_xml('<p><a href="javascript:alert(1)">x</a><script>bad()</script> ok</p>')

    assert "javascript" not in xml
    assert "bad()" not in xml
    assert "ok" in xml


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


def test_collect_page_objects_reports_image_callback_ids():
    xml = _page("<one:OE><one:Image objectID='img'><one:CallbackID callbackID='cb-1'/></one:Image></one:OE>")

    image = next(obj for obj in collect_page_objects(xml) if obj["type"] == "Image")

    assert image["callback_id"] == "cb-1"


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
