"""Daily-use edge cases found by testing against the real OneNote: each test pins one fixed bug."""

import asyncio
import xml.etree.ElementTree as ET

import pytest

from fake_onenote import NS, ONE, FakeBridge, oid
from local_onenote_mcp import app, search, server  # noqa: F401  (server registers the tools)
from local_onenote_mcp.editing import EditOperation, PageEditor
from local_onenote_mcp.hierarchy import resolve_item
from local_onenote_mcp.image_utils import ImageDimensionError, image_format
from local_onenote_mcp.pagexml import build_page_update_xml
from local_onenote_mcp.reader import render_page
from local_onenote_mcp.text_index import TextIndex
from local_onenote_mcp.tools import writing

NB, SEC, PAGE = oid(1), oid(2), oid(3)
HIERARCHY = f"""<one:Notebooks xmlns:one="{NS}">
  <one:Notebook name="NB" ID="{NB}"><one:Section name="Sec" ID="{SEC}">
    <one:Page name="Plan" ID="{PAGE}" lastModifiedTime="2026-10-05T10:00:00.000Z"/>
  </one:Section></one:Notebook>
</one:Notebooks>"""
NESTED = f"""<one:Page xmlns:one="{NS}" ID="{PAGE}">
<one:Outline objectID="o1"><one:OEChildren>
  <one:OE objectID="a"><one:T><![CDATA[parent]]></one:T><one:OEChildren>
    <one:OE objectID="a1"><one:T><![CDATA[only child]]></one:T></one:OE>
  </one:OEChildren></one:OE>
  <one:OE objectID="b"><one:T><![CDATA[2026&nbsp;plan&#160;done]]></one:T></one:OE>
</one:OEChildren></one:Outline>
</one:Page>"""


def run(tool, *args, **kwargs):
    return asyncio.run(tool(*args, **kwargs))


@pytest.fixture
def fake(monkeypatch):
    def install(pages: dict | None = None) -> FakeBridge:
        bridge = FakeBridge(HIERARCHY, {PAGE: NESTED} if pages is None else pages)
        monkeypatch.setattr(app.onenote, "bridge", bridge)
        app.onenote.invalidate()
        return bridge

    yield install
    app.onenote.invalidate()


def editor() -> PageEditor:
    return PageEditor(ET.fromstring(NESTED))


def markdown(content: str) -> str:
    return render_page(build_page_update_xml("p", content=content, content_format="markdown"), "markdown", include_title=False)


# -- writing Markdown -------------------------------------------------------


def test_numbered_lists_keep_their_start_and_restart():
    xml = build_page_update_xml("p", content="5. five\n6. six\n\nbreak\n\n1. one", content_format="markdown")

    assert xml.count("restartNumberingAt") == 2
    assert 'restartNumberingAt="5" text="5."' in xml and 'restartNumberingAt="1" text="1."' in xml


def test_a_numbered_to_do_keeps_its_number_and_its_box():
    once = markdown("1. [ ] open\n2. [x] done")

    assert once == "1. [ ] open\n2. [x] done"


def test_text_that_only_looks_like_a_tag_stays():
    assert "List\\<String>" in markdown("Use List<String> and <b>bold</b>")
    assert "**bold**" in markdown("Use List<String> and <b>bold</b>")


def test_relative_links_keep_their_target_as_text():
    assert markdown("[the doc](docs/readme.md) and <a href='javascript:alert(1)'>x</a>") == "the doc (docs/readme.md) and x"


def test_nested_quotes_and_code_in_a_quote_stay_inside_the_quote():
    content = "> outer\n>\n> > inner\n\n> code:\n>\n> ```\n> x = 1\n> ```"

    assert markdown(content) == "> outer\n>\n> > inner\n>\n> code:\n>\n> ```\n> x = 1\n> ```"


def test_two_code_blocks_in_a_row_stay_two_blocks():
    assert markdown("```\na\n```\n\n```\nb\n```") == "```\na\n```\n\n```\nb\n```"


def test_whitespace_only_content_adds_no_empty_block_and_titles_are_one_line():
    xml = build_page_update_xml("p", title="two\nlines", content="  \n ", content_format="markdown")

    assert "<one:Outline>" not in xml and "two lines" in xml


# -- edit_page ---------------------------------------------------------------


def test_set_text_reads_markdown_inline_only():
    page = editor()
    for content in ("2026.", "# not a heading", "- not a list"):
        located = page._locate("b")
        page.set_text(located, content, "markdown")
        assert ET.fromstring(f'<x xmlns:one="{NS}">{ET.tostring(located.oe, encoding="unicode")}</x>').find(f".//{ONE}T").text == content
    page.set_text(page._locate("b"), "**bold**", "markdown")
    assert page._locate("b").oe.find(ONE + "T").text == "<strong>bold</strong>"


def test_replace_text_matches_no_break_spaces():
    page = editor()

    assert page.replace_text("2026 plan done", "shipped", 0) == 1


def test_a_paragraph_under_a_deleted_one_is_gone():
    page = editor()
    page.apply(EditOperation(op="delete", paragraph_id="a"))

    with pytest.raises(ValueError, match="not on this page"):
        page.apply(EditOperation(op="set_text", paragraph_id="a1", content="x"))


def test_deleting_the_only_child_leaves_no_empty_children_list():
    page = editor()
    page.apply(EditOperation(op="delete", paragraph_id="a1"))

    assert page._locate("a").oe.find(ONE + "OEChildren") is None


# -- tools -------------------------------------------------------------------


def test_a_failed_create_page_removes_the_new_empty_page(fake, monkeypatch):
    bridge = fake()

    def refuse(xml, schema, force):
        raise RuntimeError("rejected")

    monkeypatch.setattr(bridge, "update_page_content", refuse)
    result = run(writing.create_page, SEC, "T", "x")

    assert result["ok"] is False
    assert [op for op, _ in bridge.calls if op == "delete_hierarchy"] == ["delete_hierarchy"]
    assert len(bridge.root.findall(f".//{ONE}Page")) == 1


def test_append_without_y_goes_below_the_content(fake):
    page = NESTED.replace('<one:Outline objectID="o1">', '<one:Outline objectID="o1"><one:Position x="36" y="100" z="0"/><one:Size width="50" height="40"/>')
    bridge = fake({PAGE: page})
    run(writing.append_to_page, PAGE, content="more", x=300)

    position = ET.fromstring(bridge.updates[-1]).find(f".//{ONE}Position").attrib
    assert float(position["x"]) == 300 and float(position["y"]) == 158


def test_replace_page_body_keeps_the_old_start_position(fake):
    page = NESTED.replace('<one:Outline objectID="o1">', '<one:Outline objectID="o1"><one:Position x="40" y="200" z="0"/>')
    bridge = fake({PAGE: page})
    run(writing.replace_page_body, PAGE, "new")

    position = ET.fromstring(bridge.updates[-1]).find(f".//{ONE}Position").attrib
    assert (float(position["x"]), float(position["y"])) == (40, 200)


def test_appending_text_that_renders_to_nothing_is_an_error(fake):
    bridge = fake()

    assert run(writing.append_to_page, PAGE, content="   ")["ok"] is False
    assert run(writing.append_to_page, PAGE, content="<script>x</script>", content_format="html")["ok"] is False
    assert bridge.updates == []


def test_the_text_index_forgets_a_page_this_server_changed(fake, monkeypatch):
    fake()
    index = TextIndex(None)
    monkeypatch.setattr(app.onenote, "on_page_changed", index.forget)
    index.put(PAGE, "stamp", "old text")
    run(writing.append_to_page, PAGE, content="new")

    assert index.get(PAGE, "stamp") is None


def test_image_format_comes_from_the_bytes(tmp_path):
    fake_png = tmp_path / "a.png"
    fake_png.write_text("not an image")
    gif = tmp_path / "b.png"
    gif.write_bytes(b"GIF89a" + b"\x00" * 20)

    assert image_format(gif) == "gif"
    with pytest.raises(ImageDimensionError):
        image_format(fake_png)


# -- finding things ----------------------------------------------------------


def test_paths_ignore_spaces_around_slashes_and_ambiguity_lists_ids():
    items = [
        {"type": "page", "id": "1", "name": "Note", "path": "NB/Sec/Note"},
        {"type": "page", "id": "2", "name": "Note", "path": "NB/Other/Note"},
    ]

    assert resolve_item(items, " /NB / Sec/Note/ ")["id"] == "1"
    with pytest.raises(ValueError, match=r"ID 1\).*ID 2\)"):
        resolve_item(items, "note")


def test_search_does_not_match_notebook_or_section_names():
    assert search.score(["work"], "Plan", "Work/Sec/Plan", "nothing here") == 0
    assert search.score(["plan"], "Plan", "Work/Sec/Plan", "") > 0
