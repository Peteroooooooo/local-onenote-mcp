"""Tools that drive the OneNote app itself: navigation, sync, export, closing notebooks."""

from __future__ import annotations

from pathlib import Path
from typing import Any

from ..app import WRITES, ok, onenote, tool
from ..constants import PUBLISH_FORMATS, enum_value


@tool(WRITES)
def navigate_to(target: str, paragraph_id: str = "", new_window: bool = False) -> dict[str, Any]:
    """Show a notebook, section, page or paragraph in the OneNote window. target is an identifier or a onenote: link."""

    if target.strip().casefold().startswith("onenote:"):
        onenote.call("navigate_to_url", url=target.strip(), new_window=new_window)
    else:
        object_id = onenote.resolve(target)["id"]
        onenote.call("navigate_to", object_id=object_id, page_content_object_id=paragraph_id, new_window=new_window)
    return ok(navigated=True)


@tool(WRITES)
def sync_hierarchy(object_identifier: str) -> dict[str, Any]:
    """Ask OneNote to sync a notebook, section group or section."""

    object_id = onenote.resolve(object_identifier)["id"]
    onenote.call("sync_hierarchy", object_id=object_id)
    return ok(object_id=object_id, synced=True)


@tool(WRITES)
def close_notebook(notebook_identifier: str, force: bool = False) -> dict[str, Any]:
    """Close a notebook in OneNote (its files stay on disk)."""

    notebook_id = onenote.resolve(notebook_identifier, "notebook")["id"]
    onenote.call("close_notebook", notebook_id=notebook_id, force=force)
    return ok(notebook_id=notebook_id, closed=True)


@tool(WRITES)
def publish_object(object_identifier: str, target_path: str, format: str = "pdf", overwrite: bool = False) -> dict[str, Any]:
    """Export a notebook, section or page to a file (pdf, docx, xps, mht, html, one, onepkg, emf)."""

    output = Path(target_path).expanduser().resolve()
    if output.exists() and not overwrite:
        raise ValueError(f"Target already exists: {output}")
    output.parent.mkdir(parents=True, exist_ok=True)
    object_id = onenote.resolve(object_identifier)["id"]
    result = onenote.call("publish", object_id=object_id, target_path=str(output), format=enum_value("format", format, PUBLISH_FORMATS))
    return ok(path=result["path"])
