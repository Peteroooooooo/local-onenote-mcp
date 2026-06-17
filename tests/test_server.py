import asyncio
from pathlib import Path

from local_onenote_mcp import server


def test_health_check_includes_runtime_diagnostics(monkeypatch):
    def fake_hierarchy_items(start_id, scope):
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
    assert result["search_default_backend"] == "local_scan"
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
        return [{"type": "page", "id": "page-id", "name": "Found", "path": "NB/Sec/Found"}]

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
