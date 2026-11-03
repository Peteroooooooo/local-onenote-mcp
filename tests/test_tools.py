import asyncio
import base64
import os
import xml.etree.ElementTree as ET

import pytest

from fake_onenote import NS, ONE, FakeBridge, oid
from local_onenote_mcp import app, server  # noqa: F401  (server registers the tools)
from local_onenote_mcp.tools import organizing, reading, writing

NB, SEC, PAGE = oid(1), oid(2), oid(3)
HIERARCHY = f"""<one:Notebooks xmlns:one="{NS}">
  <one:Notebook name="NB" ID="{NB}"><one:Section name="Sec" ID="{SEC}">
    <one:Page name="Plan" ID="{PAGE}" lastModifiedTime="2026-10-05T10:00:00.000Z"/>
    <one:Page name="Old" ID="{oid(4)}" lastModifiedTime="2026-09-01T10:00:00.000Z"/>
    <one:Page name="New" ID="{oid(5)}" lastModifiedTime="2026-10-20T10:00:00.000Z"/>
  </one:Section></one:Notebook>
</one:Notebooks>"""
PAGE_XML = f"""<one:Page xmlns:one="{NS}" ID="{PAGE}" lastModifiedTime="2026-10-05T10:00:00.000Z">
<one:QuickStyleDef index="0" name="p"/>
<one:Title><one:OE objectID="t"><one:T><![CDATA[Plan]]></one:T></one:OE></one:Title>
<one:Outline objectID="o1"><one:OEChildren>
  <one:OE objectID="a"><one:T><![CDATA[ship in 2025]]></one:T></one:OE>
  <one:OE objectID="b"><one:T><![CDATA[second]]></one:T></one:OE>
</one:OEChildren></one:Outline>
<one:Outline objectID="o2"><one:OEChildren><one:OE objectID="c"><one:T><![CDATA[alone]]></one:T></one:OE></one:OEChildren></one:Outline>
<one:Image objectID="img" format="png"><one:CallbackID callbackID="cb"/><one:Data>aGVsbG8=</one:Data></one:Image>
<one:InsertedFile objectID="f" preferredName="doc.pdf"><one:CallbackID callbackID="file-cb"/></one:InsertedFile>
</one:Page>"""


def run(tool, *args, **kwargs):
    return asyncio.run(tool(*args, **kwargs))


@pytest.fixture
def fake(monkeypatch):
    def install(hierarchy: str = HIERARCHY, pages: dict | None = None) -> FakeBridge:
        bridge = FakeBridge(hierarchy, {PAGE: PAGE_XML} if pages is None else pages)
        monkeypatch.setattr(app.onenote, "bridge", bridge)
        app.onenote.invalidate()
        return bridge

    yield install
    app.onenote.invalidate()


@pytest.fixture
def backups_on(monkeypatch, tmp_path):
    monkeypatch.setattr(app.backups, "path", tmp_path / "backups.sqlite3")
    monkeypatch.setattr(app.backups, "_db", None)
    monkeypatch.setattr("local_onenote_mcp.pages.data_dir", lambda: tmp_path)


def test_every_tool_is_registered_with_annotations():
    tools = {tool.name: tool for tool in app.mcp._tool_manager.list_tools()}

    assert len(tools) == 41
    assert tools["edit_page"].annotations.destructiveHint is True
    assert tools["get_page"].annotations.readOnlyHint is True
    assert tools["view_page"].annotations.readOnlyHint is True
    assert "EditOperation" in str(tools["edit_page"].parameters)


def test_failures_come_back_as_errors(fake):
    fake()
    result = run(reading.get_page, oid(999))

    assert result["ok"] is False and result["error"]


def test_get_page_lists_paragraphs(fake):
    fake()
    result = run(reading.get_page, PAGE, include_paragraphs=True, include_objects=True)

    assert [p["id"] for p in result["paragraphs"]] == ["a", "b", "c"]
    assert {obj["type"] for obj in result["objects"]} >= {"Image", "InsertedFile"}


def test_edit_page_resubmits_only_the_changed_outline(fake):
    bridge = fake()
    result = run(writing.edit_page, PAGE, [{"op": "replace_text", "find": "2025", "replace": "2026"}, {"op": "insert_after", "paragraph_id": "b", "content": "third"}])

    assert result["ok"] is True and result["results"][0]["replacements"] == 1
    assert len(bridge.updates) == 1
    update = bridge.updates[0]
    assert "ship in 2026" in update and "third" in update
    assert 'objectID="o2"' not in update and "img" not in update and "Plan" not in update
    assert bridge.deleted == []


def test_edit_page_deletes_an_outline_left_empty(fake):
    bridge = fake()
    run(writing.edit_page, PAGE, [{"op": "delete", "paragraph_id": "c"}])

    assert bridge.updates == [] and bridge.deleted == ["o2"]


def test_edit_page_validates_before_writing(fake):
    bridge = fake()
    result = run(writing.edit_page, PAGE, [{"op": "set_text", "paragraph_id": "a", "content": "x"}, {"op": "delete", "paragraph_id": "zzz"}])

    assert result["ok"] is False and bridge.updates == []


def test_delete_page_content_points_paragraphs_to_their_block(fake):
    fake()
    result = run(writing.delete_page_content, PAGE, "a")

    assert result["ok"] is False and "o1" in result["error"]


def test_whole_page_backup_restores_as_new_page_and_in_place(fake, backups_on):
    bridge = fake()
    deleted = run(writing.delete_page_content, PAGE, "img")
    listed = run(writing.list_backups, PAGE)

    assert isinstance(deleted["backup_id"], int)
    assert listed["backups"][0]["path"] == "NB/Sec/Plan" and listed["backups"][0]["kind"] == "page"

    copy = run(writing.restore_backup, deleted["backup_id"])
    restored = ET.fromstring(bridge.updates[-1])
    assert copy["ok"] is True and copy["page_id"] != PAGE
    assert restored.attrib == {"ID": copy["page_id"]}
    assert not any("objectID" in node.attrib for node in restored.iter())
    assert restored.find(f".//{ONE}CallbackID") is None
    assert restored.find(f".//{ONE}Data").text == "aGVsbG8="
    assert restored.find(f".//{ONE}InsertedFile").attrib["pathSource"].endswith("doc.pdf")

    bridge.deleted.clear()
    in_place = run(writing.restore_backup, deleted["backup_id"], mode="replace")
    assert in_place["ok"] is True
    assert bridge.calls[-len(bridge.deleted) - 1][1]["force"] is True
    assert bridge.deleted == ["o1", "o2", "img", "f"]


def test_edit_page_backs_up_only_the_changed_text_block(fake, backups_on):
    bridge = fake()
    edited = run(writing.edit_page, PAGE, [{"op": "set_text", "paragraph_id": "a", "content": "changed"}])
    listed = run(writing.list_backups, PAGE)["backups"]

    assert listed[0]["kind"] == "blocks" and listed[0]["reason"] == "edit_page"
    assert all(params.get("page_info", 0) == 0 for op, params in bridge.calls if op == "get_page_content")
    _meta, saved, _files = app.backups.load(edited["backup_id"])
    assert [node.attrib["objectID"] for node in saved.findall(ONE + "Outline")] == ["o1"]
    assert saved.find(ONE + "Image") is None and saved.find(ONE + "Title") is None
    assert run(writing.restore_backup, edited["backup_id"], mode="new_page")["ok"] is False


def test_blocks_restore_writes_the_old_outline_back(fake, backups_on):
    bridge = fake()
    edited = run(writing.edit_page, PAGE, [{"op": "set_text", "paragraph_id": "a", "content": "changed"}])
    # Since the backup: paragraph b was deleted and OneNote renumbered the style definitions.
    bridge.pages[PAGE] = (
        PAGE_XML.replace('<one:QuickStyleDef index="0" name="p"/>', '<one:QuickStyleDef index="0" name="h1"/><one:QuickStyleDef index="1" name="p"/>')
        .replace("ship in 2025", "changed")
        .replace('<one:OE objectID="b"><one:T><![CDATA[second]]></one:T></one:OE>', "")
        .replace('<one:OE objectID="a">', '<one:OE objectID="a" quickStyleIndex="1">')
    )

    restored = run(writing.restore_backup, edited["backup_id"])
    update = ET.fromstring(bridge.updates[-1])
    paragraphs = {oe.find(ONE + "T").text: oe.attrib for oe in update.iter(ONE + "OE")}

    assert restored["ok"] is True and restored["page_id"] == PAGE
    assert [node.attrib["objectID"] for node in update.findall(ONE + "Outline")] == ["o1"]
    assert update.find(ONE + "Image") is None
    assert paragraphs == {"ship in 2025": {"objectID": "a"}, "second": {}}
    assert [d.attrib["name"] for d in update.findall(ONE + "QuickStyleDef")] == ["h1", "p"]
    assert isinstance(restored["backup_id"], int)


def test_a_deleted_block_comes_back_as_a_new_outline(fake, backups_on):
    bridge = fake()
    edited = run(writing.edit_page, PAGE, [{"op": "delete", "paragraph_id": "c"}])
    bridge.pages[PAGE] = PAGE_XML.replace('<one:Outline objectID="o2"><one:OEChildren><one:OE objectID="c"><one:T><![CDATA[alone]]></one:T></one:OE></one:OEChildren></one:Outline>', "")

    run(writing.restore_backup, edited["backup_id"])
    outline = ET.fromstring(bridge.updates[-1]).find(ONE + "Outline")

    assert bridge.deleted == ["o2"]
    assert "objectID" not in outline.attrib and outline.find(f".//{ONE}T").text == "alone"


def test_restoring_blocks_remaps_style_indexes(fake, backups_on):
    bridge = fake()
    page = PAGE_XML.replace('<one:OE objectID="a">', '<one:OE objectID="a" quickStyleIndex="0">')
    bridge.pages[PAGE] = page
    edited = run(writing.edit_page, PAGE, [{"op": "set_text", "paragraph_id": "a", "content": "changed"}])
    bridge.pages[PAGE] = page.replace('<one:QuickStyleDef index="0" name="p"/>', '<one:QuickStyleDef index="0" name="h1"/><one:QuickStyleDef index="1" name="p"/>')

    run(writing.restore_backup, edited["backup_id"])
    update = ET.fromstring(bridge.updates[-1])

    assert update.find(f".//{ONE}OE[@objectID='a']").attrib["quickStyleIndex"] == "1"


def test_edit_next_to_an_image_takes_a_whole_page_backup(fake, backups_on):
    page = PAGE_XML.replace('<one:OE objectID="b">', '<one:OE objectID="pic"><one:Image format="png"><one:CallbackID callbackID="cb2"/></one:Image></one:OE><one:OE objectID="b">')
    bridge = fake(pages={PAGE: page})
    edited = run(writing.edit_page, PAGE, [{"op": "set_text", "paragraph_id": "a", "content": "changed"}])

    assert run(writing.list_backups, PAGE)["backups"][0]["kind"] == "page"
    assert any(params.get("page_info") for op, params in bridge.calls if op == "get_page_content")
    assert edited["ok"] is True


def test_transient_class_not_registered_is_retried_for_safe_calls(fake, monkeypatch):
    from local_onenote_mcp.bridge import OneNoteBridgeError
    from local_onenote_mcp.onenote import CLASS_NOT_REGISTERED

    bridge = fake()
    failures = {"get_hierarchy": 1, "update_page_content": 1}
    real_call = bridge.call

    def flaky(operation, **params):
        if failures.get(operation):
            failures[operation] -= 1
            raise OneNoteBridgeError("没有注册类", hresult=CLASS_NOT_REGISTERED)
        return real_call(operation, **params)

    monkeypatch.setattr(bridge, "call", flaky)
    monkeypatch.setattr("local_onenote_mcp.onenote.time.sleep", lambda seconds: None)

    assert run(reading.list_notebooks)["ok"] is True
    assert run(writing.update_page_title, PAGE, "x")["ok"] is False
    assert bridge.updates == []


def test_view_page_returns_one_image_per_printed_page(fake, monkeypatch):
    bridge = fake()
    published = []

    def publish(object_id, target_path, format):
        published.append(target_path)
        open(target_path, "wb").close()
        return {"path": target_path}

    def render_pdf(path, width, max_pages):
        return {"pages": 4, "images": [base64.b64encode(b"\xff\xd8jpeg").decode()] * max_pages}

    monkeypatch.setattr(bridge, "publish", publish, raising=False)
    monkeypatch.setattr(bridge, "render_pdf", render_pdf, raising=False)
    result = run(reading.view_page, PAGE, max_pages=2)

    assert result[0] == "Printed pages 1-2 of 4."
    assert [image.data for image in result[1:]] == [b"\xff\xd8jpeg"] * 2
    assert published and not os.path.exists(published[0])


def test_copy_page_writes_a_clean_copy_into_the_section(fake):
    bridge = fake()
    result = run(writing.copy_page, PAGE)

    assert result["ok"] is True and result["page"]["parent_id"] == SEC
    assert "objectID" not in bridge.updates[-1] and "<one:CallbackID" not in bridge.updates[-1]


def test_search_by_date_alone_lists_recent_pages(fake):
    fake()
    result = run(reading.search_pages, modified_after="2026-10-01")

    assert [page["name"] for page in result["pages"]] == ["New", "Plan"]
    assert result["search_backend"] == "hierarchy"
    assert run(reading.search_pages)["ok"] is False


def _section(*pages: tuple[str, int]) -> str:
    body = "".join(f'<one:Page ID="{page_id}" name="{page_id}" pageLevel="{level}"/>' for page_id, level in pages)
    return f'<one:Notebooks xmlns:one="{NS}"><one:Notebook ID="NB" name="NB"><one:Section ID="S" name="Sec">{body}</one:Section></one:Notebook></one:Notebooks>'


def _order(tree: ET.Element) -> list[tuple[str, str]]:
    return [(node.attrib["ID"], node.attrib["pageLevel"]) for node in tree if node.tag == ONE + "Page"]


def test_move_page_to_top_carries_its_subpages(fake):
    bridge = fake(_section(("A", 1), ("B", 1), ("B1", 2), ("B2", 3), ("C", 1)), {})
    result = run(organizing.move_page, "B", to_top=True)

    assert _order(bridge.submitted[0]) == [("B", "1"), ("B1", "2"), ("B2", "3"), ("A", "1"), ("C", "1")]
    assert [entry["new_id"] for entry in result["moved_pages"]] == ["B", "B1", "B2"]


def test_move_page_after_a_page_skips_that_pages_subpages(fake):
    bridge = fake(_section(("A", 1), ("A1", 2), ("B", 1), ("C", 1)), {})
    run(organizing.move_page, "C", after_page_identifier="A")

    assert _order(bridge.submitted[0]) == [("A", "1"), ("A1", "2"), ("C", "1"), ("B", "1")]


def test_moving_a_subpage_alone_promotes_it(fake):
    bridge = fake(_section(("A", 1), ("A1", 2), ("B", 1)), {})
    run(organizing.move_page, "A1", include_subpages=False)

    assert _order(bridge.submitted[0]) == [("A", "1"), ("B", "1"), ("A1", "1")]


def test_move_page_rejects_its_own_subpage_as_anchor(fake):
    fake(_section(("A", 1), ("A1", 2), ("B", 1)), {})
    result = run(organizing.move_page, "A", after_page_identifier="A1")

    assert result["ok"] is False and "cannot be the page being moved" in result["error"]


def test_set_page_level_needs_a_parent_page(fake):
    bridge = fake(_section(("A", 1), ("B", 1)), {})

    assert run(organizing.set_page_level, "A", 2)["ok"] is False
    assert run(organizing.set_page_level, "B", 3)["ok"] is False
    assert run(organizing.set_page_level, "B", 2)["ok"] is True
    assert _order(bridge.submitted[0]) == [("A", "1"), ("B", "2")]


def test_move_section_refuses_name_clash_and_moving_into_itself(fake):
    bridge = fake(
        f'<one:Notebooks xmlns:one="{NS}">'
        '<one:Notebook ID="NB" name="NB"><one:Section ID="S" name="Sec"/>'
        '<one:SectionGroup ID="G" name="G"><one:SectionGroup ID="I" name="Inner"/></one:SectionGroup></one:Notebook>'
        '<one:Notebook ID="NB2" name="Other"><one:Section ID="X" name="sec"/></one:Notebook></one:Notebooks>',
        {},
    )

    clash = run(organizing.move_section, "NB/Sec", "Other")
    itself = run(organizing.move_section, "NB/G", "NB/G/Inner")

    assert clash["ok"] is False and "already has" in clash["error"]
    assert itself["ok"] is False and "into itself" in itself["error"]
    assert bridge.submitted == []


def test_delete_hierarchy_refuses_notebooks(fake):
    fake()
    result = run(organizing.delete_hierarchy, NB)

    assert result["ok"] is False and "close_notebook" in result["error"]
