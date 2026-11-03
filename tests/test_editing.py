import xml.etree.ElementTree as ET

import pytest

from local_onenote_mcp.editing import EditOperation, PageEditor

NS = "http://schemas.microsoft.com/office/onenote/2013/onenote"
ONE = f"{{{NS}}}"

PAGE = f"""<one:Page xmlns:one="{NS}" ID="P">
<one:TagDef index="0" type="0" symbol="3" fontColor="automatic" highlightColor="none" name="To Do"/>
<one:QuickStyleDef index="0" name="PageTitle" font="Calibri Light" fontSize="20.0"/>
<one:QuickStyleDef index="1" name="p" font="Calibri" fontSize="11.0"/>
<one:Title><one:OE objectID="t"><one:T><![CDATA[Title]]></one:T></one:OE></one:Title>
<one:Outline objectID="o1"><one:OEChildren>
  <one:OE objectID="a" quickStyleIndex="1"><one:T><![CDATA[Due <span style='font-weight:bold'>2025</span> plan 2025]]></one:T></one:OE>
  <one:OE objectID="b"><one:List><one:Bullet bullet="2"/></one:List><one:T><![CDATA[bullet]]></one:T>
    <one:OEChildren><one:OE objectID="b1"><one:T><![CDATA[child]]></one:T></one:OE></one:OEChildren></one:OE>
  <one:OE objectID="c"><one:Tag index="0" completed="false"/><one:T><![CDATA[task]]></one:T></one:OE>
</one:OEChildren></one:Outline>
<one:Outline objectID="o2"><one:OEChildren><one:OE objectID="d"><one:T><![CDATA[alone]]></one:T></one:OE></one:OEChildren></one:Outline>
<one:Image objectID="img"><one:CallbackID callbackID="cb"/></one:Image>
</one:Page>"""


def editor(xml: str = PAGE) -> PageEditor:
    return PageEditor(ET.fromstring(xml))


def run(page: PageEditor, **op) -> dict:
    return page.apply(EditOperation(**op))


def tags(update: str) -> dict[str, dict]:
    """Tag attributes by the objectID of their paragraph."""

    return {oe.attrib["objectID"]: oe.find(ONE + "Tag").attrib for oe in ET.fromstring(update).iter(ONE + "OE") if oe.find(ONE + "Tag") is not None and "objectID" in oe.attrib}


def test_paragraphs_list_ids_depth_lists_and_todos():
    paragraphs = editor().paragraphs()

    assert [p["id"] for p in paragraphs] == ["a", "b", "b1", "c", "d"]
    assert paragraphs[0]["text"] == "Due 2025 plan 2025"
    assert paragraphs[1]["list"] == "bullet" and paragraphs[2]["depth"] == 1
    assert paragraphs[3]["todo"] is False


def test_replace_text_keeps_formatting_and_respects_max_count():
    page = editor()
    assert run(page, op="replace_text", find="2025", replace="2026")["replacements"] == 2
    assert page.page.find(f".//{ONE}OE[@objectID='a']/{ONE}T").text == "Due <span style='font-weight:bold'>2026</span> plan 2026"

    once = editor()
    run(once, op="replace_text", find="2025", replace="X", max_count=1)
    assert once.page.find(f".//{ONE}OE[@objectID='a']/{ONE}T").text.count("X") == 1


def test_replace_text_reports_missing_text():
    with pytest.raises(ValueError, match="not found"):
        run(editor(), op="replace_text", find="nowhere")


def test_replace_text_matches_across_font_runs():
    # OneNote stores mixed Chinese/Latin text as separate font runs.
    split = PAGE.replace(
        "Due <span style='font-weight:bold'>2025</span> plan 2025",
        "<span style='font-family:\"Microsoft YaHei\"'>每分钟</span><span lang=en-US> 60 </span>"
        "<span style='font-family:\"Microsoft YaHei\"'>次。</span>a<br>b",
    )
    page = editor(split)

    assert run(page, op="replace_text", find="60 次", replace="120 次")["replacements"] == 1
    text = page.page.find(f".//{ONE}OE[@objectID='a']/{ONE}T").text
    assert text == "<span style='font-family:\"Microsoft YaHei\"'>每分钟</span><span lang=en-US> 120 次</span><span style='font-family:\"Microsoft YaHei\"'>。</span>a<br>b"
    with pytest.raises(ValueError, match="not found"):
        run(page, op="replace_text", find="ab", replace="x")


def test_set_text_keeps_bullet_and_children():
    page = editor()
    run(page, op="set_text", paragraph_id="b", content="**new** text")
    node = page.page.find(f".//{ONE}OE[@objectID='b']")

    assert node.find(ONE + "T").text == "<strong>new</strong> text"
    assert node.find(ONE + "List") is not None and node.find(ONE + "OEChildren") is not None


def test_insert_after_reuses_page_definitions():
    page = editor()
    added = run(page, op="insert_after", paragraph_id="b", content="## New\n\n- [ ] item")
    update = page.build_update("P")

    assert added["inserted_paragraphs"] == 2
    assert [p.get("text") for p in page.paragraphs()][:5] == ["Due 2025 plan 2025", "bullet", "child", "New", "item"]
    assert '<one:QuickStyleDef index="2" name="h2"' in update
    assert '<one:OE quickStyleIndex="2">' in update
    item = [oe for oe in ET.fromstring(update).iter(ONE + "OE") if oe.find(ONE + "T").text == "item"][0]
    assert item.find(ONE + "Tag").attrib == {"index": "0", "completed": "false"}
    assert update.count("<one:TagDef") == 1


def test_update_holds_only_changed_outlines_with_definitions_first():
    page = editor()
    run(page, op="set_todo", paragraph_id="c", done=True)
    update = page.build_update("P")
    root = ET.fromstring(update)

    assert [o.attrib["objectID"] for o in root.findall(ONE + "Outline")] == ["o1"]
    assert root.find(ONE + "Image") is None and root.find(ONE + "Title") is None
    assert [child.tag.removeprefix(ONE) for child in root][:4] == ["TagDef", "QuickStyleDef", "QuickStyleDef", "Outline"]
    assert tags(update)["c"]["completed"] == "true"


def test_deleting_the_last_paragraph_deletes_its_outline():
    page = editor()
    run(page, op="delete", paragraph_id="d")

    assert page.empty_outlines() == ["o2"]
    assert page.build_update("P") is None
    with pytest.raises(ValueError, match="not on this page"):
        run(page, op="set_text", paragraph_id="d", content="x")


def test_set_todo_adds_a_tag_definition_when_the_page_has_none():
    page = editor(PAGE.replace('<one:TagDef index="0" type="0" symbol="3" fontColor="automatic" highlightColor="none" name="To Do"/>', ""))
    run(page, op="set_todo", paragraph_id="a", done=True)
    update = page.build_update("P")

    assert '<one:TagDef index="0" type="0" symbol="3"' in update
    assert tags(update)["a"] == {"index": "0", "completed": "true"}


def test_set_text_refuses_non_text_paragraphs():
    xml = PAGE.replace("<one:T><![CDATA[alone]]></one:T>", "<one:Image/>")
    with pytest.raises(ValueError, match="not text"):
        run(editor(xml), op="set_text", paragraph_id="d", content="x")


TABLE_PAGE = f"""<one:Page xmlns:one="{NS}" ID="P">
<one:QuickStyleDef index="0" name="p"/>
<one:Outline objectID="o1"><one:OEChildren><one:OE objectID="tbl"><one:Table>
  <one:Columns><one:Column index="0" width="60"/><one:Column index="1" width="60"/></one:Columns>
  <one:Row objectID="r0"><one:Cell objectID="c00" shadingColor="#DEEAF6"><one:OEChildren><one:OE objectID="h1"><one:T><![CDATA[Name]]></one:T></one:OE></one:OEChildren></one:Cell>
    <one:Cell objectID="c01"><one:OEChildren><one:OE objectID="h2"><one:T><![CDATA[State]]></one:T></one:OE></one:OEChildren></one:Cell></one:Row>
  <one:Row objectID="r1"><one:Cell objectID="c10"><one:OEChildren><one:OE objectID="x1"><one:T><![CDATA[Alpha]]></one:T></one:OE></one:OEChildren></one:Cell>
    <one:Cell objectID="c11"><one:OEChildren><one:OE objectID="x2"><one:T><![CDATA[Open]]></one:T></one:OE><one:OE objectID="x3"><one:T><![CDATA[note]]></one:T></one:OE></one:OEChildren></one:Cell></one:Row>
</one:Table></one:OE></one:OEChildren></one:Outline>
</one:Page>"""


def rows(page: PageEditor) -> list[list[str]]:
    table = page.page.find(f".//{ONE}Table")
    return [[" / ".join(t.text for t in cell.iter(ONE + "T")) for cell in row.findall(ONE + "Cell")] for row in table.findall(ONE + "Row")]


def test_table_cells_are_paragraphs_with_their_position():
    paragraphs = editor(TABLE_PAGE).paragraphs()

    assert paragraphs[0] == {"id": "tbl", "depth": 0, "kind": "table (2 rows x 2 columns)"}
    assert {p["id"]: (p["row"], p["column"]) for p in paragraphs[1:]} == {"h1": (0, 0), "h2": (0, 1), "x1": (1, 0), "x2": (1, 1), "x3": (1, 1)}
    assert all(p["table"] == "tbl" for p in paragraphs[1:])


def test_set_text_in_a_cell_and_add_rows():
    page = editor(TABLE_PAGE)
    run(page, op="set_text", paragraph_id="x2", content="Done")
    run(page, op="add_row", paragraph_id="h1", cells=["Beta", "**New**"])
    run(page, op="add_row", paragraph_id="tbl", cells=["Gamma"])

    assert rows(page) == [["Name", "State"], ["Beta", "<strong>New</strong>"], ["Alpha", "Done / note"], ["Gamma", ""]]
    added = page.page.find(f".//{ONE}Table").findall(ONE + "Row")[1]
    assert not any("objectID" in node.attrib for node in added.iter())
    assert ET.fromstring(page.build_update("P")).find(ONE + "Outline").attrib["objectID"] == "o1"


def test_add_row_refuses_too_many_cells_and_non_table_paragraphs():
    with pytest.raises(ValueError, match="2 columns"):
        run(editor(TABLE_PAGE), op="add_row", paragraph_id="x1", cells=["a", "b", "c"])
    with pytest.raises(ValueError, match="inside a table"):
        run(editor(), op="add_row", paragraph_id="a", cells=["a"])


def test_delete_row_and_cell_paragraphs():
    page = editor(TABLE_PAGE)
    run(page, op="delete", paragraph_id="x3")
    run(page, op="delete", paragraph_id="x1")
    assert rows(page) == [["Name", "State"], ["", "Open"]]

    run(page, op="delete_row", paragraph_id="x2")
    assert rows(page) == [["Name", "State"]]
    with pytest.raises(ValueError, match="not on this page"):
        run(page, op="set_text", paragraph_id="x2", content="x")
    with pytest.raises(ValueError, match="only row"):
        run(page, op="delete_row", paragraph_id="h1")