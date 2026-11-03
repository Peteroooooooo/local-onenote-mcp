import asyncio
import xml.etree.ElementTree as ET

from local_onenote_mcp import server

NS = "http://schemas.microsoft.com/office/onenote/2013/onenote"


def _section(*pages: tuple[str, int]) -> str:
    body = "".join(f'<one:Page ID="{page_id}" name="{page_id}" pageLevel="{level}"/>' for page_id, level in pages)
    return f'<one:Section xmlns:one="{NS}" ID="S" name="Sec">{body}</one:Section>'


def _order(tree: ET.Element) -> list[tuple[str, str]]:
    return [(node.attrib["ID"], node.attrib["pageLevel"]) for node in tree if node.tag.endswith("}Page")]


def _setup(monkeypatch, xml: str):
    submitted = []
    item = {"type": "page", "id": "B", "path": "NB/Sec/B", "parent_id": "S"}
    monkeypatch.setattr(server, "_resolve_item", lambda identifier, item_type=None: dict(item, id=identifier, path=f"NB/Sec/{identifier}"))
    monkeypatch.setattr(server, "_resolve_id", lambda identifier, item_type=None: identifier)
    monkeypatch.setattr(server, "_tree", lambda start_id, scope: ET.fromstring(xml))
    monkeypatch.setattr(server, "_submit_tree", lambda tree: submitted.append(tree))
    monkeypatch.setattr(server, "_lookup_by_id", lambda object_id: {"id": object_id})
    return submitted


def test_move_page_to_top_carries_its_subpages(monkeypatch):
    submitted = _setup(monkeypatch, _section(("A", 1), ("B", 1), ("B1", 2), ("B2", 3), ("C", 1)))

    result = asyncio.run(server.move_page("B", to_top=True))

    assert result["ok"] is True
    assert _order(submitted[0]) == [("B", "1"), ("B1", "2"), ("B2", "3"), ("A", "1"), ("C", "1")]
    assert [entry["new_id"] for entry in result["moved_pages"]] == ["B", "B1", "B2"]


def test_move_page_after_a_page_skips_that_pages_subpages(monkeypatch):
    submitted = _setup(monkeypatch, _section(("A", 1), ("A1", 2), ("B", 1), ("C", 1)))

    asyncio.run(server.move_page("C", after_page_identifier="A"))

    assert _order(submitted[0]) == [("A", "1"), ("A1", "2"), ("C", "1"), ("B", "1")]


def test_moving_a_subpage_alone_promotes_it(monkeypatch):
    submitted = _setup(monkeypatch, _section(("A", 1), ("A1", 2), ("B", 1)))

    asyncio.run(server.move_page("A1", include_subpages=False))

    assert _order(submitted[0]) == [("A", "1"), ("B", "1"), ("A1", "1")]


def test_move_page_rejects_its_own_subpage_as_anchor(monkeypatch):
    _setup(monkeypatch, _section(("A", 1), ("A1", 2), ("B", 1)))

    result = asyncio.run(server.move_page("A", after_page_identifier="A1"))

    assert result["ok"] is False
    assert "cannot be the page being moved" in result["error"]


def test_set_page_level_needs_a_parent_page(monkeypatch):
    _setup(monkeypatch, _section(("A", 1), ("B", 1)))

    first = asyncio.run(server.set_page_level("A", 2))
    third = asyncio.run(server.set_page_level("B", 3))

    assert first["ok"] is False
    assert third["ok"] is False


def test_move_section_refuses_a_name_clash(monkeypatch):
    items = {
        "Sec": {"type": "section", "id": "S", "name": "Sec", "path": "NB/Sec", "parent_id": "NB"},
        "Other": {"type": "notebook", "id": "NB2", "name": "Other", "path": "Other", "parent_id": None},
    }
    target = f'<one:Notebook xmlns:one="{NS}" ID="NB2" name="Other"><one:Section ID="X" name="sec"/></one:Notebook>'
    monkeypatch.setattr(server, "_resolve_item", lambda identifier, item_type=None: items[identifier])
    monkeypatch.setattr(server, "_tree", lambda start_id, scope: ET.fromstring(target))
    monkeypatch.setattr(server, "_submit_tree", lambda tree: (_ for _ in ()).throw(AssertionError("must not submit")))

    result = asyncio.run(server.move_section("Sec", "Other"))

    assert result["ok"] is False
    assert "already contains" in result["error"]


def test_move_section_group_into_itself_is_rejected(monkeypatch):
    items = {
        "G": {"type": "section_group", "id": "G", "name": "G", "path": "NB/G", "parent_id": "NB"},
        "G/Inner": {"type": "section_group", "id": "I", "name": "Inner", "path": "NB/G/Inner", "parent_id": "G"},
    }
    monkeypatch.setattr(server, "_resolve_item", lambda identifier, item_type=None: items[identifier])

    result = asyncio.run(server.move_section("G", "G/Inner"))

    assert result["ok"] is False
    assert "into itself" in result["error"]
