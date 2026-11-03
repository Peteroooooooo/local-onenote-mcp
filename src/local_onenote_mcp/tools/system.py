"""Tools that drive the OneNote app itself: navigation, sync, export, closing notebooks."""

from __future__ import annotations

from pathlib import Path
from typing import Any

from ..app import WRITES, ok, onenote, tool
from ..constants import FILING_LOCATION_TYPES, FILING_LOCATIONS, PUBLISH_FORMATS, enum_value


@tool(WRITES)
def navigate_to(object_identifier: str, page_content_object_id: str = "", new_window: bool = False) -> dict[str, Any]:
    """Show a notebook, section, page or paragraph in the OneNote window."""

    object_id = onenote.resolve(object_identifier)["id"]
    onenote.call("navigate_to", object_id=object_id, page_content_object_id=page_content_object_id, new_window=new_window)
    return ok(navigated=True)


@tool(WRITES)
def navigate_to_url(url: str, new_window: bool = False) -> dict[str, Any]:
    """Open a onenote: link in the OneNote window."""

    onenote.call("navigate_to_url", url=url, new_window=new_window)
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


@tool(WRITES)
def set_filing_location(filing_location: str, filing_location_type: str, section_or_page_identifier: str) -> dict[str, Any]:
    """Choose where OneNote files emails, web clips, printouts and similar content."""

    object_id = onenote.resolve(section_or_page_identifier)["id"]
    onenote.call(
        "set_filing_location",
        filing_location=enum_value("filing_location", filing_location, FILING_LOCATIONS),
        filing_location_type=enum_value("filing_location_type", filing_location_type, FILING_LOCATION_TYPES),
        section_or_page_id=object_id,
    )
    return ok(object_id=object_id, updated=True)
