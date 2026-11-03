"""End-to-end check against the real OneNote desktop app.

Creates a throwaway notebook in a temporary folder, exercises the write/read/
search/image/delete tools through the MCP stdio transport, exports the main
test page as PDF for visual inspection, and closes the notebook again. Your
existing notebooks are only read (search runs across them).

    .venv\\Scripts\\python.exe scripts\\e2e_onenote.py
"""

from __future__ import annotations

import asyncio
import json
import os
import struct
import sys
import tempfile
import time
import zlib
from pathlib import Path
from typing import Any

from mcp import ClientSession, StdioServerParameters
from local_onenote_mcp.hierarchy import parse_xml
from local_onenote_mcp.pagexml import content_bottom
from mcp.client.stdio import stdio_client

TOKEN = f"E2E_TOKEN_{int(time.time())}"
MARKDOWN = f"""# Weekly plan

Intro with **bold**, *italic*, `code` and a [link](https://example.com).

## Lists

- first bullet
- second bullet
  - nested bullet
1. step one
2. step two

- [ ] open task
- [x] finished task

> quoted line

| Item | Owner | Status |
| --- | --- | --- |
| Beta | Infra | **In progress** |
| Cutover | Ops | Pending |

```python
def f():
    return 1
```

{TOKEN}
"""


def tiny_png(path: Path, width: int = 48, height: int = 24) -> None:
    rows = b"".join(b"\x00" + bytes([200, 60, 60]) * width for _ in range(height))

    def chunk(kind: bytes, data: bytes) -> bytes:
        return struct.pack(">I", len(data)) + kind + data + struct.pack(">I", zlib.crc32(kind + data) & 0xFFFFFFFF)

    header = struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0)
    path.write_bytes(b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", header) + chunk(b"IDAT", zlib.compress(rows)) + chunk(b"IEND", b""))


class Checks:
    def __init__(self) -> None:
        self.rows: list[dict[str, Any]] = []

    def add(self, name: str, ok: bool, detail: Any = None) -> None:
        self.rows.append({"check": name, "ok": bool(ok), **({"detail": detail} if detail is not None else {})})


async def main() -> int:
    root = Path(tempfile.mkdtemp(prefix="onenote-mcp-e2e-"))
    image = root / "red.png"
    tiny_png(image)
    params = StdioServerParameters(
        command=sys.executable,
        args=["-m", "local_onenote_mcp.server"],
        env={**os.environ, "LOCAL_ONENOTE_MCP_BACKGROUND_INDEX": "0", "LOCAL_ONENOTE_MCP_BACKUP_PATH": str(root / "backups.sqlite3")},
    )
    checks = Checks()
    timings: dict[str, int] = {}

    async with stdio_client(params) as (read, write):
        async with ClientSession(read, write) as session:
            await session.initialize()

            async def call(name: str, **args: Any) -> Any:
                started = time.perf_counter()
                result = await session.call_tool(name, args)
                timings[name] = round((time.perf_counter() - started) * 1000)
                if result.isError:
                    return {"ok": False, "error": result.content[0].text if result.content else ""}
                if result.content and result.content[0].type == "image":
                    return {"ok": True, "image": result.content[0]}
                return json.loads(result.content[0].text)

            notebook = await call("create_notebook", name_or_path=str(root / "MCP E2E"))
            notebook_id = notebook["notebook_id"]
            try:
                section = await call("create_section", parent_identifier=notebook_id, section_name="Scratch")
                section_id = section["section_id"]

                first = await call("create_page", section_identifier=section_id, title="Same title", content=MARKDOWN, content_format="markdown")
                second = await call("create_page", section_identifier=section_id, title="Same title", content="second body")
                checks.add("create_page returns distinct IDs for duplicate titles", first["page_id"] != second["page_id"])

                page = await call("get_page", page_identifier=first["page_id"], format="markdown", include_objects=True, include_xml=True)
                text = page.get("text", "")
                xml = page.get("xml", "")
                checks.add("headings use native OneNote styles", 'name="h1"' in xml and 'name="h2"' in xml)
                checks.add("to-dos are native To Do tags", 'name="To Do"' in xml and 'completed="true"' in xml)
                checks.add("markdown read-back keeps structure", all(
                    piece in text for piece in ("# Weekly plan\n", "\n## Lists\n", "- first bullet", "  - nested bullet", "2. step two", "- [x] finished task", "| Beta | Infra |", "```")
                ), text)

                before = await call("get_page", page_identifier=first["page_id"], include_xml=True)
                await call("add_image_to_page", page_identifier=first["page_id"], image_path=str(image), width=96)
                after = await call("get_page", page_identifier=first["page_id"], include_objects=True, include_xml=True)
                objects = after["objects"]

                image_y = None
                for node in parse_xml(after["xml"]).iter():
                    if node.tag.endswith("}Outline") and node.find(".//{*}Image") is not None:
                        position = node.find("{*}Position")
                        image_y = float(position.attrib["y"]) if position is not None else None
                text_bottom = content_bottom(before["xml"]) or 0
                checks.add("new image is placed below existing text", image_y is not None and image_y >= text_bottom, {"image_y": image_y, "text_bottom": text_bottom})
                callback = next((obj.get("callback_id") for obj in objects if obj["type"] == "Image" and obj.get("callback_id")), None)
                shown = await call("get_page_image", page_identifier=first["page_id"], callback_id=callback or "")
                checks.add("get_page_image returns an image", shown.get("ok") and "image" in shown, shown.get("error"))

                found = await call("search_pages", query=TOKEN, start_identifier=notebook_id)
                checks.add("search finds the new page", any(p["id"] == first["page_id"] for p in found.get("pages", [])), found.get("search_backend"))

                missing = await call("get_page", page_identifier="{00000000-0000-0000-0000-000000000000}{1}{B0}")
                checks.add("bad identifiers come back as MCP errors", missing.get("ok") is False, missing.get("error"))

                pdf = root / "weekly-plan.pdf"
                exported = await call("publish_object", object_identifier=first["page_id"], target_path=str(pdf), format="pdf", overwrite=True)
                checks.add("publish_object writes a PDF", exported.get("ok") and pdf.exists() and pdf.stat().st_size > 1000, str(pdf))

                await edit_checks(call, checks, root, section_id, first["page_id"])

                await call("replace_page_body", page_identifier=second["page_id"], content="replaced body")
                replaced = await call("get_page", page_identifier=second["page_id"])
                checks.add("replace_page_body swaps the body", replaced.get("text") == "replaced body", replaced.get("text"))

                gone = await call("delete_hierarchy", object_identifier=second["page_id"], permanently=True)
                remaining = [p["id"] for p in (await call("list_pages", section_identifier=section_id)).get("pages", [])]
                checks.add("deleting one duplicate keeps the other", first["page_id"] in remaining and second["page_id"] not in remaining)
                back = await call("restore_backup", backup_id=(gone.get("backup_ids") or [0])[0])
                back_text = (await call("get_page", page_identifier=back.get("page_id", ""))).get("text")
                checks.add("a permanently deleted page can be restored from its backup", back.get("ok") and back_text == "replaced body", back.get("error") or back_text)
                await call("delete_hierarchy", object_identifier=back.get("page_id", ""), permanently=True)

                await move_checks(call, checks, root, notebook_id, section_id)
            finally:
                for notebook in ("MCP E2E", "MCP E2E Target"):
                    if (root / notebook).exists():
                        await call("close_notebook", notebook_identifier=notebook, force=True)

    failed = [row for row in checks.rows if not row["ok"]]
    print(json.dumps({"ok": not failed, "folder": str(root), "checks": checks.rows, "timings_ms": timings}, ensure_ascii=False, indent=2))
    return 0 if not failed else 1


async def edit_checks(call, checks: Checks, root: Path, section_id: str, page_id: str) -> None:
    async def read(pid: str) -> dict[str, Any]:
        return await call("get_page", page_identifier=pid, format="markdown", include_paragraphs=True, include_objects=True)

    def images(page: dict[str, Any]) -> int:
        return sum(1 for obj in page.get("objects", []) if obj["type"] == "Image")

    before = await read(page_id)
    ids = {p.get("text"): p["id"] for p in before["paragraphs"]}
    edited = await call(
        "edit_page",
        page_identifier=page_id,
        operations=[
            {"op": "replace_text", "find": "step two", "replace": "step 2"},
            {"op": "set_text", "paragraph_id": ids["second bullet"], "content": "**edited** bullet"},
            {"op": "insert_after", "paragraph_id": ids["first bullet"], "content": "- inserted bullet"},
            {"op": "set_todo", "paragraph_id": ids["open task"], "done": True},
            {"op": "delete", "paragraph_id": ids["quoted line"]},
        ],
    )
    after = await read(page_id)
    text = after.get("text", "")
    checks.add(
        "edit_page changes only the targeted paragraphs",
        edited.get("ok")
        and all(piece in text for piece in ("2. step 2", "- **edited** bullet", "- first bullet\n- inserted bullet", "- [x] open task", "| Beta | Infra |", "```", TOKEN))
        and "quoted line" not in text,
        edited.get("error") or text,
    )
    checks.add("edit_page keeps images and the title", images(after) == images(before) == 1 and after.get("title") == before.get("title"), {"before": images(before), "after": images(after), "title": after.get("title")})

    listed = await call("list_backups", page_identifier=page_id)
    checks.add("edit_page saved a backup", edited.get("backup_id") in [b["backup_id"] for b in listed.get("backups", [])], listed)

    restored = await call("restore_backup", backup_id=edited.get("backup_id") or 0, mode="replace")
    back = await read(page_id)
    checks.add(
        "restore_backup(replace) brings the page back exactly",
        restored.get("ok") and back.get("text") == before.get("text") and images(back) == 1 and back.get("title") == before.get("title"),
        restored.get("error") or {"title": back.get("title"), "images": images(back), "text": back.get("text")},
    )

    attachment = root / "notes.txt"
    attachment.write_text("attachment body ✓", encoding="utf-8")
    attached = await call("add_file_to_page", page_identifier=page_id, file_path=str(attachment))
    original = await read(page_id)
    copied = await call("copy_page", page_identifier=page_id)
    copy = await read(copied.get("page_id", ""))
    files = [obj for obj in copy.get("objects", []) if obj["type"] == "InsertedFile"]
    saved = root / "copied-notes.txt"
    if files:
        await call("get_binary_content", page_identifier=copied["page_id"], file_name=files[0].get("preferredName", ""), save_to=str(saved))
    checks.add(
        "copy_page copies text, images and attachments",
        attached.get("ok") and copied.get("ok") and copy.get("text") == original.get("text") and images(copy) == 1
        and saved.exists() and saved.read_bytes() == attachment.read_bytes(),
        attached.get("error") or copied.get("error") or {"files": files, "images": images(copy), "copy": copy.get("text"), "original": original.get("text")},
    )

    today = time.strftime("%Y-%m-%d")
    recent = await call("search_pages", modified_after=today, start_identifier=section_id, max_results=50)
    checks.add("search_pages filters by modification date", page_id in [p["id"] for p in recent.get("pages", [])], recent.get("error") or recent.get("count"))


async def move_checks(call, checks: Checks, root: Path, notebook_id: str, section_id: str) -> None:
    async def page_order(section: str) -> list[tuple[str, str]]:
        pages = (await call("list_pages", section_identifier=section)).get("pages", [])
        return [(p["name"], p.get("pageLevel", "1")) for p in pages]

    parent = await call("create_page", section_identifier=section_id, title="Parent page", content="parent body")
    child = await call("create_page", section_identifier=section_id, title="Child page", content="child body")
    await call("create_page", section_identifier=section_id, title="Other page")
    indented = await call("set_page_level", page_identifier=child["page_id"], level=2)
    checks.add("set_page_level makes a subpage", indented.get("ok"), indented.get("error"))

    reordered = await call("move_page", page_identifier="MCP E2E/Scratch/Other page", to_top=True)
    order = await page_order(section_id)
    checks.add("move_page reorders inside a section", reordered.get("ok") and order[0][0] == "Other page", order)

    target_nb = (await call("create_notebook", name_or_path=str(root / "MCP E2E Target")))["notebook_id"]
    dest = (await call("create_section", parent_identifier=target_nb, section_name="Dest"))["section_id"]
    moved = await call("move_page", page_identifier=parent["page_id"], destination_section_identifier=dest)
    dest_order = await page_order(dest)
    body = await call("get_page", page_identifier=moved.get("page_id", ""))
    checks.add(
        "move_page moves a page with its subpage to another notebook",
        moved.get("ok") and dest_order == [("Parent page", "1"), ("Child page", "2")] and body.get("text") == "parent body",
        {"result": moved.get("error") or moved.get("moved_pages"), "dest": dest_order},
    )
    checks.add("moved pages are gone from the source", all(name not in {"Parent page", "Child page"} for name, _ in await page_order(section_id)))

    await call("create_section", parent_identifier=notebook_id, section_name="Movable")
    await call("create_section_group", parent_identifier=notebook_id, group_name="Group")
    into_group = await call("move_section", section_identifier="MCP E2E/Movable", destination_identifier="MCP E2E/Group")
    checks.add("move_section moves a section into a section group", into_group.get("ok") and into_group["item"]["path"] == "MCP E2E/Group/Movable", into_group.get("error"))

    across = await call("move_section", section_identifier="MCP E2E/Group/Movable", destination_identifier=target_nb)
    checks.add("move_section moves a section to another notebook", across.get("ok") and across["item"]["path"] == "MCP E2E Target/Movable", across.get("error") or across.get("warning"))

    renamed = await call("rename_section", section_identifier="MCP E2E Target/Movable", new_name="Renamed")
    checks.add("rename_section renames", renamed.get("ok") and renamed["item"]["path"] == "MCP E2E Target/Renamed", renamed.get("error"))

    await call("close_notebook", notebook_identifier=notebook_id)
    await call("close_notebook", notebook_identifier=target_nb)
    reopened = [
        await call("open_hierarchy", path=str(root / "MCP E2E"), create_type="notebook"),
        await call("open_hierarchy", path=str(root / "MCP E2E Target"), create_type="notebook"),
    ]
    time.sleep(2)
    all_sections = await call("list_sections")
    sections = sorted(s["path"] for s in all_sections.get("sections", []) if s["path"].startswith("MCP E2E"))
    expected = ["MCP E2E Target/Dest", "MCP E2E Target/Renamed", "MCP E2E/Scratch"]
    checks.add(
        "moves and renames survive closing and reopening the notebooks",
        sections == expected,
        {"sections": sections, "reopen": [r.get("error") or r.get("object_id") for r in reopened], "list_error": all_sections.get("error"),
         "notebooks": [n["path"] for n in (await call("list_notebooks")).get("notebooks", [])]},
    )
    dest_after = await page_order("MCP E2E Target/Dest")
    checks.add("moved pages survive reopening", dest_after == [("Parent page", "1"), ("Child page", "2")], dest_after)


if __name__ == "__main__":
    raise SystemExit(asyncio.run(main()))
