import asyncio
import base64
from pathlib import Path

from local_onenote_mcp import server


def test_health_check_includes_runtime_diagnostics(monkeypatch):
    def fake_hierarchy_items(start_id, scope, fresh=False):
        assert start_id == ""
        assert scope == "sections"
        return [
            {"type": "notebook", "name": "NB", "path": "NB"},
            {"type": "section", "name": "Sec", "path": "NB/Sec"},
        ]

    monkeypatch.setattr(server, "_hierarchy_items", fake_hierarchy_items)

    result = asyncio.run(server.health_check())

    assert result["ok"] is True
    assert result["server"] == "local-onenote"
    assert result["identifier_resolution_order"] == ["id", "exact_path", "unique_name"]
    assert result["search_default_backend"] == "auto"
    assert result["content_formats"] == ["plain", "html", "markdown"]
    assert result["python_executable"]
    assert result["module_path"].endswith("server.py")


def test_resolve_identifier_returns_single_item(monkeypatch):
    expected = {"type": "section", "id": "section-id", "path": "NB/Sec"}

    def fake_resolve_item(identifier, item_type=None):
        assert identifier == "NB/Sec"
        assert item_type == "section"
        return expected

    monkeypatch.setattr(server, "_resolve_item", fake_resolve_item)

    result = asyncio.run(server.resolve_identifier("NB/Sec", "section"))

    assert result["ok"] is True
    assert result["item"] == expected
    assert result["identifier_resolution_order"] == ["id", "exact_path", "unique_name"]


def test_resolve_identifier_rejects_unknown_type():
    result = asyncio.run(server.resolve_identifier("NB/Sec", "folder"))

    assert result["ok"] is False
    assert "item_type must be empty or one of" in result["error"]


def test_without_recycle_bin_removes_container_and_children():
    items = [
        {"name": "Notebook", "path": "Notebook"},
        {"name": "OneNote_RecycleBin", "path": "Notebook/OneNote_RecycleBin", "isRecycleBin": "true"},
        {"name": "Deleted", "path": "Notebook/OneNote_RecycleBin/Deleted", "isInRecycleBin": "true"},
        {"name": "Active", "path": "Notebook/Section/Active"},
    ]

    filtered = server._without_recycle_bin(items)

    assert filtered == [
        {"name": "Notebook", "path": "Notebook"},
        {"name": "Active", "path": "Notebook/Section/Active"},
    ]


def test_search_pages_include_unindexed_uses_local_scan(monkeypatch):
    bridge_called = False

    def fake_bridge(*args, **kwargs):
        nonlocal bridge_called
        bridge_called = True
        raise AssertionError("OneNote index should not be used for include_unindexed=True")

    def fake_local_text_search(start_id, query, max_results, include_recycle_bin):
        assert start_id == ""
        assert query == "needle"
        assert max_results == 3
        assert include_recycle_bin is False
        pages = [{"type": "page", "id": "page-id", "name": "Found", "path": "NB/Sec/Found"}]
        return {"pages": pages, "total_matches": 1, "scanned": 5, "total_pages": 5, "scan_complete": True}

    monkeypatch.setattr(server, "_bridge", fake_bridge)
    monkeypatch.setattr(server, "_local_text_search", fake_local_text_search)

    result = asyncio.run(
        server.search_pages(
            "needle",
            max_results=3,
            include_snippets=False,
            include_unindexed=True,
        )
    )

    assert result["ok"] is True
    assert result["count"] == 1
    assert result["search_backend"] == "local_scan"
    assert result["scan_complete"] is True
    assert bridge_called is False


def test_publish_object_resolves_target_path_before_bridge(monkeypatch, tmp_path):
    captured = {}

    def fake_resolve_id(identifier, item_type=None):
        assert identifier == "page-id"
        assert item_type is None
        return "resolved-page-id"

    def fake_bridge(operation, **params):
        captured["operation"] = operation
        captured["params"] = params
        return {"path": params["target_path"]}

    monkeypatch.chdir(tmp_path)
    monkeypatch.setattr(server, "_resolve_id", fake_resolve_id)
    monkeypatch.setattr(server, "_bridge", fake_bridge)

    result = asyncio.run(
        server.publish_object(
            "page-id",
            "exports/out.pdf",
            format="pdf",
            overwrite=True,
        )
    )

    expected = tmp_path / "exports" / "out.pdf"
    assert result["ok"] is True
    assert captured["operation"] == "publish"
    assert Path(captured["params"]["target_path"]) == expected
    assert result["path"] == str(expected)


def test_open_hierarchy_resolves_existing_friendly_path_without_bridge(monkeypatch):
    bridge_called = False
    expected = {"type": "section", "id": "section-id", "path": "Notebook/Group/Sec", "name": "Sec"}

    def fake_find_item_by_path(path, item_type=None):
        assert path == "Notebook/Group/Sec"
        assert item_type is None
        return expected

    def fake_bridge(*args, **kwargs):
        nonlocal bridge_called
        bridge_called = True
        raise AssertionError("Existing friendly paths should resolve without OpenHierarchy")

    monkeypatch.setattr(server, "_find_item_by_path", fake_find_item_by_path)
    monkeypatch.setattr(server, "_bridge", fake_bridge)

    result = asyncio.run(server.open_hierarchy("Notebook/Group/Sec"))

    assert result["ok"] is True
    assert result["object_id"] == "section-id"
    assert result["item"] == expected
    assert result["opened_existing"] is True
    assert bridge_called is False


def test_create_section_returns_refreshed_current_section_id(monkeypatch):
    parent = {"type": "section_group", "id": "group-id", "path": "Notebook/Group", "name": "Group"}
    refreshed = {"type": "section", "id": "current-section-id", "path": "Notebook/Group/New Sec", "name": "New Sec"}

    monkeypatch.setattr(server, "_resolve_item", lambda identifier, item_type=None: parent)
    monkeypatch.setattr(server, "_bridge", lambda operation, **params: {"object_id": "stale-section-id"})
    monkeypatch.setattr(
        server,
        "_refresh_created_item",
        lambda **kwargs: refreshed,
    )

    result = asyncio.run(server.create_section("Notebook/Group", "New Sec"))

    assert result["ok"] is True
    assert result["section_id"] == "current-section-id"
    assert result["section"] == refreshed
    assert result["path"] == "Notebook/Group/New Sec"


def test_delete_page_content_rejects_non_deletable_child_with_parent_suggestion(monkeypatch):
    page_xml = """<one:Page xmlns:one="http://schemas.microsoft.com/office/onenote/2013/onenote" ID="p">
    <one:Outline objectID="outline-id"><one:OEChildren><one:OE objectID="oe-id">
      <one:T><![CDATA[hello]]></one:T>
    </one:OE></one:OEChildren></one:Outline>
    </one:Page>"""

    monkeypatch.setattr(server, "_resolve_id", lambda identifier, item_type=None: "page-id")
    monkeypatch.setattr(server, "_page_xml", lambda page_id, page_info="basic": page_xml)

    result = asyncio.run(server.delete_page_content("page-id", "oe-id"))

    assert result["ok"] is False
    assert "not directly deletable" in result["error"]
    assert "outline-id" in result["error"]


def test_delete_hierarchy_retries_when_same_path_reappears_with_new_id(monkeypatch):
    calls = []
    initial = {"type": "section_group", "id": "old-id", "path": "Notebook/Test", "name": "Test"}
    remaining = {"type": "section_group", "id": "new-id", "path": "Notebook/Test", "name": "Test"}

    def fake_bridge(operation, **params):
        calls.append(params["object_id"])
        return {"deleted": True}

    def fake_find_item_by_path(path, item_type=None):
        assert path == "Notebook/Test"
        assert item_type == "section_group"
        return remaining if len(calls) == 1 else None

    monkeypatch.setattr(server, "_resolve_item", lambda identifier, item_type=None: initial)
    monkeypatch.setattr(server, "_bridge", fake_bridge)
    monkeypatch.setattr(server, "_find_item_by_path", fake_find_item_by_path)
    monkeypatch.setattr(server.time, "sleep", lambda seconds: None)

    result = asyncio.run(server.delete_hierarchy("Notebook/Test", permanently=True))

    assert result["ok"] is True
    assert result["deleted_ids"] == ["old-id", "new-id"]
    assert result["verified_gone"] is True


PAGE = {"type": "page", "id": "page-a", "name": "Untitled", "path": "NB/Sec/Untitled", "lastModifiedTime": "t1"}
SIBLING = {"type": "page", "id": "page-b", "name": "Untitled", "path": "NB/Sec/Untitled", "lastModifiedTime": "t1"}


def _raise(message):
    def fail(*args, **kwargs):
        raise AssertionError(message)

    return fail


def test_delete_page_never_touches_a_same_titled_sibling(monkeypatch):
    deleted = []

    def fake_bridge(operation, **params):
        assert operation == "delete_hierarchy"
        deleted.append(params["object_id"])
        return {"deleted": True}

    monkeypatch.setattr(server, "_resolve_item", lambda identifier, item_type=None: dict(PAGE))
    monkeypatch.setattr(server, "_bridge", fake_bridge)
    monkeypatch.setattr(server, "_hierarchy_items", lambda start_id="", scope="pages", fresh=False: [dict(SIBLING)])
    monkeypatch.setattr(server.time, "sleep", lambda seconds: None)

    result = asyncio.run(server.delete_hierarchy("page-a"))

    assert result["ok"] is True
    assert deleted == ["page-a"]


def test_create_page_returns_new_id_even_when_title_already_exists(monkeypatch):
    section = {"type": "section", "id": "sec", "path": "NB/Sec", "name": "Sec"}
    new_page = dict(SIBLING, id="page-new")

    def fake_bridge(operation, **params):
        if operation == "create_new_page":
            return {"page_id": "page-new"}
        return {"updated": True}

    monkeypatch.setattr(server, "_resolve_item", lambda identifier, item_type=None: section)
    monkeypatch.setattr(server, "_bridge", fake_bridge)
    monkeypatch.setattr(
        server, "_hierarchy_items", lambda start_id="", scope="pages", fresh=False: [dict(PAGE), new_page]
    )

    result = asyncio.run(server.create_page("NB/Sec", "Untitled", "body"))

    assert result["ok"] is True
    assert result["page_id"] == "page-new"
    assert result["page"]["id"] == "page-new"


def test_replace_page_body_writes_new_content_before_deleting_old(monkeypatch):
    page_xml = """<one:Page xmlns:one="http://schemas.microsoft.com/office/onenote/2013/onenote" ID="p">
    <one:Outline objectID="old-outline"><one:OEChildren><one:OE><one:T><![CDATA[old]]></one:T></one:OE></one:OEChildren></one:Outline>
    </one:Page>"""
    calls = []

    def fake_bridge(operation, **params):
        calls.append(operation)
        if operation == "update_page_content":
            raise RuntimeError("OneNote rejected the XML")
        return {}

    monkeypatch.setattr(server, "_resolve_id", lambda identifier, item_type=None: "p")
    monkeypatch.setattr(server, "_page_xml", lambda page_id, page_info="basic": page_xml)
    monkeypatch.setattr(server, "_bridge", fake_bridge)

    result = asyncio.run(server.replace_page_body("p", "new"))

    assert result["ok"] is False
    assert calls == ["update_page_content"]


def test_search_auto_falls_back_to_scan_when_index_finds_nothing(monkeypatch):
    monkeypatch.setattr(server, "_index_coverage", lambda start_id, include_recycle_bin: (10, 400))
    monkeypatch.setattr(server, "_index_search", lambda start_id, query: [])
    monkeypatch.setattr(
        server,
        "_local_text_search",
        lambda start_id, query, limit, include_recycle_bin: {
            "pages": [dict(PAGE)],
            "total_matches": 1,
            "scanned": 10,
            "total_pages": 40,
            "scan_complete": False,
        },
    )

    result = asyncio.run(server.search_pages("needle", include_snippets=False))

    assert result["search_backend"] == "local_scan"
    assert result["scan_complete"] is False
    assert result["total_pages"] == 40


def test_search_auto_prefers_index_hits_while_text_index_is_incomplete(monkeypatch):
    monkeypatch.setattr(server, "_index_coverage", lambda start_id, include_recycle_bin: (10, 400))
    monkeypatch.setattr(server, "_index_search", lambda start_id, query: [dict(PAGE)])
    monkeypatch.setattr(server, "_local_text_search", _raise("index hits should not trigger a scan"))

    result = asyncio.run(server.search_pages("needle", include_snippets=False))

    assert result["search_backend"] == "onenote_index"
    assert [page["id"] for page in result["pages"]] == ["page-a"]


def test_local_scan_matches_all_terms_and_caches_page_text(monkeypatch):
    reads = []

    def fake_page_xml(page_id, page_info="basic"):
        reads.append(page_id)
        return (
            '<one:Page xmlns:one="http://schemas.microsoft.com/office/onenote/2013/onenote" ID="p">'
            "<one:Outline><one:OEChildren><one:OE><one:T><![CDATA[alpha beta]]></one:T></one:OE>"
            "</one:OEChildren></one:Outline></one:Page>"
        )

    server.text_index.clear()
    monkeypatch.setattr(server, "_hierarchy_items", lambda start_id="", scope="pages", fresh=False: [dict(PAGE)])
    monkeypatch.setattr(server, "_page_xml", fake_page_xml)

    hit = server._local_text_search("", "beta alpha", 5, False)
    miss = server._local_text_search("", "alpha gamma", 5, False)

    assert [page["id"] for page in hit["pages"]] == ["page-a"]
    assert miss["pages"] == []
    assert miss["scan_complete"] is True
    assert reads == ["page-a"]


def test_local_scan_ranks_title_hits_first(monkeypatch):
    pages = [
        {"type": "page", "id": "body", "name": "Notes", "path": "NB/S/Notes", "lastModifiedTime": "t"},
        {"type": "page", "id": "title", "name": "Kalman filter", "path": "NB/S/Kalman filter", "lastModifiedTime": "t"},
    ]
    server.text_index.clear()
    server.text_index.put("body", "t", "the kalman gain")
    server.text_index.put("title", "t", "intro")
    monkeypatch.setattr(server, "_hierarchy_items", lambda start_id="", scope="pages", fresh=False: [dict(p) for p in pages])

    result = server._local_text_search("", "kalman", 5, False)

    assert [page["id"] for page in result["pages"]] == ["title", "body"]
    assert result["total_matches"] == 2


def test_search_auto_uses_complete_text_index_without_asking_onenote(monkeypatch):
    server.text_index.clear()
    server.text_index.put("page-a", "t1", "needle in text")
    monkeypatch.setattr(server, "_hierarchy_items", lambda start_id="", scope="pages", fresh=False: [dict(PAGE)])
    monkeypatch.setattr(server, "_index_search", _raise("a complete text index should answer by itself"))

    result = asyncio.run(server.search_pages("needle"))

    assert result["search_backend"] == "text_index"
    assert result["pages"][0]["snippet"] == "needle in text"


def test_search_auto_rereads_a_few_edited_pages_instead_of_leaving_the_index(monkeypatch):
    edited = dict(PAGE, lastModifiedTime="t2")
    server.text_index.clear()
    server.text_index.put("page-a", "t1", "old text")
    monkeypatch.setattr(server, "_hierarchy_items", lambda start_id="", scope="pages", fresh=False: [dict(edited)])
    monkeypatch.setattr(server, "_index_search", _raise("a nearly complete text index should answer by itself"))
    monkeypatch.setattr(
        server,
        "_page_xml",
        lambda page_id, page_info="basic": (
            '<one:Page xmlns:one="http://schemas.microsoft.com/office/onenote/2013/onenote" ID="p">'
            "<one:Outline><one:OEChildren><one:OE><one:T><![CDATA[fresh needle]]></one:T></one:OE>"
            "</one:OEChildren></one:Outline></one:Page>"
        ),
    )

    result = asyncio.run(server.search_pages("needle", include_snippets=False))

    assert result["search_backend"] == "text_index"
    assert [page["id"] for page in result["pages"]] == ["page-a"]


def test_background_indexer_does_not_restart_on_every_hierarchy_read(monkeypatch):
    started = []
    monkeypatch.setattr(server, "BACKGROUND_INDEX", True)
    indexer = server.BackgroundIndexer()
    monkeypatch.setattr(indexer, "_run", lambda: started.append(1))

    indexer.kick()
    indexer._thread.join()
    indexer.kick()
    indexer._thread.join()
    indexer.kick(force=True)
    indexer._thread.join()

    assert started == [1, 1]


def test_text_index_persists_and_detects_stale_entries(tmp_path):
    from local_onenote_mcp.text_index import TextIndex

    path = tmp_path / "index.sqlite3"
    first = TextIndex(path)
    first.put("p1", "stamp-1", "hello")
    first.put("p2", "stamp-1", "bye")

    second = TextIndex(path)
    assert second.get("p1", "stamp-1") == "hello"
    assert second.get("p1", "stamp-2") is None
    assert second.prune({"p1"}) == 1
    assert TextIndex(path).get("p2", "stamp-1") is None


def test_get_page_image_returns_image_content(monkeypatch):
    png = b"\x89PNG\r\n\x1a\n" + b"\x00" * 16
    monkeypatch.setattr(server, "_resolve_id", lambda identifier, item_type=None: "p")
    monkeypatch.setattr(server, "_bridge", lambda operation, **params: {"base64": base64.b64encode(png).decode()})

    result = asyncio.run(server.get_page_image("p", "cb"))

    assert result.to_image_content().mimeType == "image/png"


def test_tool_failures_reach_mcp_clients_as_errors(monkeypatch):
    monkeypatch.setattr(server, "_resolve_item", lambda identifier, item_type=None: (_ for _ in ()).throw(ValueError("No page found")))

    async def call():
        return await server.mcp._tool_manager.call_tool("get_page", {"page_identifier": "missing"})

    try:
        asyncio.run(call())
    except Exception as exc:
        assert "No page found" in str(exc)
    else:
        raise AssertionError("expected a tool error")


def test_resolve_id_uses_direct_lookup_for_onenote_ids(monkeypatch):
    object_id = "{88EFFC10-08B0-4069-A2F1-918FAB9B9676}{1}{E19540698918910900793320119759617745046030661}"
    monkeypatch.setattr(server, "_lookup_by_id", lambda value: {"type": "page", "id": value})
    monkeypatch.setattr(server, "_resolve_item", _raise("an ID should not need a hierarchy scan"))

    assert server._resolve_id(object_id, "page") == object_id
    try:
        server._resolve_id(object_id, "section")
    except ValueError as exc:
        assert "expected a section" in str(exc)
    else:
        raise AssertionError("type mismatch was not reported")


def test_write_operations_invalidate_hierarchy_cache(monkeypatch):
    def fake_call(operation, **params):
        return {"xml": "<x/>"} if operation == "get_hierarchy" else {}

    monkeypatch.setattr(server.bridge, "call", fake_call)
    server._invalidate_hierarchy_cache()
    server._cached_hierarchy("", "sections", fresh=False)
    assert server._hierarchy_cache

    server._bridge("update_page_content", xml="<x/>", schema=2, force=False)

    assert not server._hierarchy_cache


def test_tools_publish_read_only_and_destructive_annotations():
    tools = {tool.name: tool for tool in asyncio.run(server.mcp.list_tools())}

    assert tools["get_page"].annotations.readOnlyHint is True
    assert tools["delete_hierarchy"].annotations.destructiveHint is True
    assert tools["create_page"].annotations.destructiveHint is False
