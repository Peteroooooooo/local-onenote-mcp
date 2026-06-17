from local_onenote_mcp.image_utils import image_dimensions, proportional_dimensions
from local_onenote_mcp import xml_utils
from local_onenote_mcp.xml_utils import (
    build_image_page_update_xml,
    build_page_update_xml,
    collect_page_objects,
    normalize_content,
    parse_hierarchy,
    resolve_item,
    text_from_page_xml,
)


def test_plain_text_becomes_onenote_inline_html():
    assert normalize_content("a<b\nc", "plain") == "a&lt;b<br/>c"


def test_html_sanitizer_keeps_safe_inline_tags():
    assert normalize_content("<p>Hello <strong>world</strong></p><script>x</script>", "html") == (
        "Hello <strong>world</strong>"
    )


def test_html_sanitizer_maps_daily_inline_styles():
    html = normalize_content("<p><s>gone</s> <code>x=1</code> <mark>note</mark></p>", "html")

    assert "text-decoration:line-through" in html
    assert "font-family:Consolas" in html
    assert "background:#FFF2CC" in html


def test_build_page_update_xml_uses_cdata():
    xml = build_page_update_xml("page-id", title="Title", content="Hello\nWorld")
    assert 'ID="page-id"' in xml
    assert "<one:Title>" in xml
    assert "<![CDATA[Hello]]>" in xml
    assert "<![CDATA[World]]>" in xml


def test_html_table_becomes_native_onenote_table():
    xml = build_page_update_xml(
        "page-id",
        title="Title",
        content="<p>Before</p><table><tr><th>Due</th><th>Task</th></tr><tr><td>2026-07-15</td><td>Submit documents</td></tr></table><p>After</p>",
        content_format="html",
    )

    assert "<one:Table" in xml
    assert 'bordersVisible="true"' in xml
    assert 'hasHeaderRow="false"' in xml
    assert '<one:OE alignment="left"><one:Table' in xml
    assert '<one:Column index="0"' in xml
    assert 'shadingColor="#D9EAF7"' in xml
    assert 'quickStyleIndex="0"' in xml
    assert "<![CDATA[<span style='font-weight:bold'>Due</span>]]>" in xml
    assert "<![CDATA[Submit documents]]>" in xml
    assert "<![CDATA[Before]]>" in xml
    assert "<![CDATA[After]]>" in xml


def test_html_table_cells_are_padded_to_column_count():
    xml = build_page_update_xml(
        "page-id",
        content="<table><tr><th>A</th><th>B</th><th>C</th></tr><tr><td>1</td><td>2</td></tr></table>",
        content_format="html",
    )

    assert xml.count("<one:Column index=") == 3
    assert xml.count("<one:Cell") == 6


def test_html_table_cells_preserve_inline_formatting():
    xml = build_page_update_xml(
        "page-id",
        content='<table><tr><td><strong>Bold</strong> <s>gone</s> <a href="https://example.com">link</a></td></tr></table>',
        content_format="html",
    )

    assert "<strong>Bold</strong>" in xml
    assert "text-decoration:line-through" in xml
    assert 'href="https://example.com"' in xml


def test_markdown_content_uses_onemore_markdig_html(monkeypatch):
    monkeypatch.setattr(
        xml_utils,
        "markdown_to_html",
        lambda content: "<h1>Heading</h1><table><thead><tr><th>A</th><th>B</th></tr></thead><tbody><tr><td>1</td><td><strong>2</strong></td></tr></tbody></table>",
    )

    xml = build_page_update_xml("page-id", content="# Heading", content_format="markdown")

    assert "<![CDATA[<span style=\"font-size:20.0pt;font-weight:bold\">Heading</span>]]>" in xml
    assert "<one:Table" in xml
    assert "<![CDATA[<strong>2</strong>]]>" in xml


def test_build_image_xml_omits_size_when_dimensions_are_missing():
    xml = build_image_page_update_xml("page-id", image_base64="abc", image_format="png")
    assert "<one:Image" in xml
    assert "<one:Size" not in xml


def test_build_image_xml_includes_size_when_dimensions_are_complete():
    xml = build_image_page_update_xml(
        "page-id",
        image_base64="abc",
        image_format="png",
        width=320,
        height=180,
    )
    assert '<one:Size width="320.00" height="180.00"/>' in xml


def test_png_dimensions_support_proportional_image_sizing(tmp_path):
    image_path = tmp_path / "sample.png"
    image_path.write_bytes(
        b"\x89PNG\r\n\x1a\n"
        b"\x00\x00\x00\rIHDR"
        b"\x00\x00\x02\x00"
        b"\x00\x00\x01\x00"
        b"\x08\x02\x00\x00\x00"
        b"\x00\x00\x00\x00"
    )

    assert image_dimensions(image_path) == (512, 256)
    assert proportional_dimensions(image_path, width=256, height=None) == (256.0, 128.0)


def test_text_from_page_xml_extracts_inline_html():
    xml = """<one:Page xmlns:one="http://schemas.microsoft.com/office/onenote/2013/onenote" ID="p">
    <one:Outline><one:OEChildren><one:OE><one:T><![CDATA[Hello<br/>World]]></one:T></one:OE></one:OEChildren></one:Outline>
    </one:Page>"""
    assert text_from_page_xml(xml) == "Hello\nWorld"


def test_collect_page_objects_keeps_idless_images_with_container():
    xml = """<one:Page xmlns:one="http://schemas.microsoft.com/office/onenote/2013/onenote" ID="p">
    <one:Outline objectID="outline-id"><one:OEChildren><one:OE objectID="oe-id">
      <one:Image format="png"><one:Data>abc</one:Data></one:Image>
    </one:OE></one:OEChildren></one:Outline>
    </one:Page>"""

    objects = collect_page_objects(xml)

    image = next(obj for obj in objects if obj["type"] == "Image")
    assert image["container_object_id"] == "oe-id"
    assert image["format"] == "png"


def test_collect_page_objects_marks_deletable_containers_and_child_suggestions():
    xml = """<one:Page xmlns:one="http://schemas.microsoft.com/office/onenote/2013/onenote" ID="p">
    <one:Outline objectID="outline-id"><one:OEChildren><one:OE objectID="oe-id">
      <one:T><![CDATA[hello]]></one:T>
    </one:OE></one:OEChildren></one:Outline>
    </one:Page>"""

    objects = collect_page_objects(xml)

    outline = next(obj for obj in objects if obj["type"] == "Outline")
    oe = next(obj for obj in objects if obj["type"] == "OE")
    assert outline["delete_supported"] is True
    assert outline["delete_object_id"] == "outline-id"
    assert oe["delete_supported"] is False
    assert oe["delete_object_id"] == "outline-id"


def test_parse_hierarchy_flattens_paths():
    xml = """<one:Notebooks xmlns:one="http://schemas.microsoft.com/office/onenote/2013/onenote">
      <one:Notebook name="NB" ID="n">
        <one:SectionGroup name="Group" ID="g">
          <one:Section name="Sec" ID="s">
            <one:Page name="Page" ID="p" />
          </one:Section>
        </one:SectionGroup>
      </one:Notebook>
    </one:Notebooks>"""
    items = parse_hierarchy(xml)
    assert [item["type"] for item in items] == ["notebook", "section_group", "section", "page"]
    assert items[-1]["path"] == "NB/Group/Sec/Page"


def test_resolve_item_prefers_exact_path_before_same_name():
    xml = """<one:Notebooks xmlns:one="http://schemas.microsoft.com/office/onenote/2013/onenote">
      <one:Notebook name="Notebook" ID="notebook-id" />
      <one:Notebook name="Projects" ID="projects-id">
        <one:Section name="People" ID="section-id">
          <one:Page name="notebook" ID="page-id" />
        </one:Section>
      </one:Notebook>
    </one:Notebooks>"""

    item = resolve_item(parse_hierarchy(xml), "Notebook")

    assert item["type"] == "notebook"
    assert item["id"] == "notebook-id"
