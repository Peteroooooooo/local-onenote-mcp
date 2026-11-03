"""An in-memory stand-in for the PowerShell OneNote bridge, used by the tool tests."""

from __future__ import annotations

import base64
import copy
import itertools
import xml.etree.ElementTree as ET

NS = "http://schemas.microsoft.com/office/onenote/2013/onenote"
ONE = f"{{{NS}}}"
ET.register_namespace("one", NS)
_ids = itertools.count(100)


def oid(n: int | str) -> str:
    """A OneNote-shaped object ID."""

    return "{%08d-0000-0000-0000-000000000000}{1}{B0}" % int(n)


class FakeBridge:
    def __init__(self, hierarchy: str, pages: dict[str, str] | None = None) -> None:
        self.root = ET.fromstring(hierarchy)
        self.pages = dict(pages or {})
        self.calls: list[tuple[str, dict]] = []
        self.updates: list[str] = []
        self.deleted: list[str] = []
        self.submitted: list[ET.Element] = []

    def call(self, operation: str, **params):
        self.calls.append((operation, params))
        return getattr(self, operation)(**params)

    def close(self) -> None:
        pass

    def _node(self, object_id: str) -> ET.Element:
        for node in self.root.iter():
            if node.attrib.get("ID") == object_id:
                return node
        raise RuntimeError(f"object not found: {object_id}")

    def get_hierarchy(self, start_id: str, scope: int, schema: int):
        node = copy.deepcopy(self._node(start_id) if start_id else self.root)
        if scope == 0:
            for child in list(node):
                node.remove(child)
        elif scope < 4:
            for parent in list(node.iter()):
                for child in list(parent):
                    if child.tag == ONE + "Page":
                        parent.remove(child)
        return {"xml": ET.tostring(node, encoding="unicode")}

    def get_page_content(self, page_id: str, page_info: int, schema: int):
        if page_id not in self.pages:
            raise RuntimeError(f"page not found: {page_id}")
        return {"xml": self.pages[page_id]}

    def update_page_content(self, xml: str, schema: int, force: bool):
        self.updates.append(xml)
        return {}

    def delete_page_content(self, page_id: str, object_id: str, force: bool):
        self.deleted.append(object_id)
        return {}

    def update_hierarchy(self, xml: str, schema: int):
        self.submitted.append(ET.fromstring(xml))
        return {}

    def create_new_page(self, section_id: str, new_page_style: int):
        page_id = oid(next(_ids))
        ET.SubElement(self._node(section_id), ONE + "Page", {"ID": page_id, "name": "", "pageLevel": "1"})
        self.pages[page_id] = f'<one:Page xmlns:one="{NS}" ID="{page_id}"/>'
        return {"page_id": page_id}

    def delete_hierarchy(self, object_id: str, permanently: bool):
        node = self._node(object_id)
        for parent in self.root.iter():
            if node in list(parent):
                parent.remove(node)
        return {}

    def get_binary_page_content(self, page_id: str, callback_id: str):
        return {"base64": base64.b64encode(f"bytes of {callback_id}".encode()).decode()}

    def get_special_location(self, location: int):
        return {"path": f"C:/special/{location}"}
