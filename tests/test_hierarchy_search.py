from datetime import datetime, timezone

import pytest

from local_onenote_mcp import search
from local_onenote_mcp.hierarchy import child_path, parse_hierarchy, resolve_item, safe_leaf_name, without_recycle_bin
from local_onenote_mcp.text_index import TextIndex

NS = 'xmlns:one="http://schemas.microsoft.com/office/onenote/2013/onenote"'
TREE = f"""<one:Notebooks {NS}>
  <one:Notebook name="NB" ID="n">
    <one:SectionGroup name="Group" ID="g"><one:Section name="Sec" ID="s" path="C:/nb/Group/Sec.one">
      <one:Page name="Page" ID="p1"/><one:Page name="Twin" ID="p2"/>
    </one:Section></one:SectionGroup>
    <one:Section name="Other" ID="o"><one:Page name="Twin" ID="p3"/></one:Section>
    <one:SectionGroup name="OneNote_RecycleBin" ID="bin" isRecycleBin="true">
      <one:Section name="Deleted" ID="d"><one:Page name="Page" ID="p4" isInRecycleBin="true"/></one:Section>
    </one:SectionGroup>
  </one:Notebook>
</one:Notebooks>"""


def test_parse_hierarchy_flattens_paths_and_parents():
    items = parse_hierarchy(TREE)
    page = next(item for item in items if item["id"] == "p1")

    assert page["path"] == "NB/Group/Sec/Page"
    assert page["parent_id"] == "s" and page["notebook_name"] == "NB" and page["section_name"] == "Sec"
    assert next(item for item in items if item["id"] == "s")["onenote_path"] == "C:/nb/Group/Sec.one"


def test_resolve_item_prefers_live_items_and_reports_ambiguity():
    items = parse_hierarchy(TREE)

    assert resolve_item(items, "Page", "page")["id"] == "p1"
    assert resolve_item(items, "nb/other/twin")["id"] == "p3"
    with pytest.raises(ValueError, match="Ambiguous"):
        resolve_item(items, "Twin", "page")
    with pytest.raises(ValueError, match="No section"):
        resolve_item(items, "Nope", "section")


def test_recycle_bin_is_filtered_with_its_contents():
    live = {item["id"] for item in without_recycle_bin(parse_hierarchy(TREE))}

    assert {"bin", "d", "p4"}.isdisjoint(live) and "p1" in live


def test_names_and_child_paths():
    assert safe_leaf_name(' a/b:c  "d" ') == "a b c d"
    assert child_path("NB", "Sec.one") == "NB/Sec"
    with pytest.raises(ValueError):
        safe_leaf_name("///")


def test_score_needs_every_term_and_ranks_titles_first():
    assert search.score(["alpha", "beta"], "Alpha", "NB/S/Alpha", "beta") > search.score(["alpha", "beta"], "x", "NB/S/x", "alpha beta")
    assert search.score(["alpha", "gamma"], "Alpha", "", "beta") == 0


def test_date_window_uses_page_modification_time():
    after = search.parse_when("2026-10-01T00:00Z")
    before = search.parse_when("2026-10-02T00:00:00+00:00")
    inside = {"lastModifiedTime": "2026-10-01T12:00:00.000Z"}
    outside = {"lastModifiedTime": "2026-10-03T00:00:00.000Z"}

    assert after == datetime(2026, 10, 1, tzinfo=timezone.utc)
    assert search.in_window(inside, after, before) and not search.in_window(outside, after, before)
    assert not search.in_window({}, after, None) and search.in_window({}, None, None)
    assert search.parse_when("2026-10-01").tzinfo is not None
    with pytest.raises(ValueError, match="ISO date"):
        search.parse_when("last week")


class _Pages:
    def __init__(self, texts):
        self.texts = texts
        self.reads = 0

    def page_xml(self, page_id):
        self.reads += 1
        body = self.texts[page_id]
        return f'<one:Page {NS} ID="{page_id}"><one:Outline><one:OEChildren><one:OE><one:T><![CDATA[{body}]]></one:T></one:OE></one:OEChildren></one:Outline></one:Page>'


def test_scan_reads_missing_pages_and_caches_them():
    pages = [{"id": "a", "name": "A", "path": "NB/S/A", "lastModifiedTime": "1"}, {"id": "b", "name": "B", "path": "NB/S/B", "lastModifiedTime": "1"}]
    onenote, index = _Pages({"a": "needle here", "b": "nothing"}), TextIndex(None)

    first = search.scan(onenote, index, [dict(p) for p in pages], ["needle"], limit=5, budget_seconds=30)
    second = search.scan(onenote, index, [dict(p) for p in pages], ["needle"], limit=5, budget_seconds=30)

    assert [p["id"] for p in first["pages"]] == ["a"] and first["scan_complete"]
    assert onenote.reads == 2 and second["total_matches"] == 1
    assert search.coverage(index, pages) == 0
    assert search.snippet("x " * 200 + "needle", ["needle"]).endswith("needle")


def test_text_index_detects_stale_entries(tmp_path):
    index = TextIndex(tmp_path / "index.sqlite3")
    index.put("a", "t1", "text")

    assert TextIndex(tmp_path / "index.sqlite3").get("a", "t1") == "text"
    assert index.get("a", "t2") is None
    assert index.prune({"b"}) == 1 and len(index) == 0
