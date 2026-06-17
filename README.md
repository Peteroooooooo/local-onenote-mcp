# Local OneNote MCP

A pure-local Microsoft OneNote MCP server for Windows. It controls the OneNote
desktop app through the local OneNote COM API, so it does not need Azure,
Microsoft Graph, API keys, or OAuth.

## Design

- Local only: every operation happens through the installed OneNote desktop app.
- COM-first: no direct binary `.one` editing.
- Safe bridge: user input is passed through JSON temp files, never interpolated
  into PowerShell script text.
- Rich surface: hierarchy, page XML/text, search, create, update, export,
  navigation, sync, and advanced raw XML tools.

PowerShell is used only as a fixed COM bridge because some Windows/Office
installations expose OneNote COM to PowerShell while leaving the COM type
library unavailable to Python automation libraries.

## Requirements

- Windows
- Microsoft OneNote desktop app
- Python 3.11+
- `uv` or `pip`

## Install

```powershell
cd local-onenote-mcp
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -e .
```

## Codex MCP Config

Add an MCP server entry that runs the installed module:

```json
{
  "mcpServers": {
    "local-onenote": {
      "command": "C:\\path\\to\\local-onenote-mcp\\.venv\\Scripts\\python.exe",
      "args": [
        "-m",
        "local_onenote_mcp.server"
      ],
      "env": {
        "LOCAL_ONENOTE_MCP_TIMEOUT": "90",
        "LOCAL_ONENOTE_MCP_MAX_TEXT_CHARS": "60000"
      }
    }
  }
}
```

Restart the MCP client after changing its config.

Make sure `command` points to an existing Python executable inside the installed
repository venv. If the path is stale, the MCP client will not expose the
`local-onenote` tools.

Validate the Codex config from this checkout:

```powershell
.\.venv\Scripts\python.exe scripts\check_codex_config.py
```

The check fails if Codex points at a missing Python executable or imports
`local_onenote_mcp` from a different checkout.

If OneMore is installed, the server can use OneMore's bundled Markdig parser
for Markdown writes. The default path is detected from the OneMore registry
entry. Override it with `LOCAL_ONENOTE_MARKDIG_DLL` when needed.

## Tools

Read and discovery:

- `health_check`
- `resolve_identifier`
- `get_special_locations`
- `list_hierarchy`
- `list_notebooks`
- `list_sections`
- `list_pages`
- `get_page`
- `get_page_xml`
- `get_page_text`
- `get_page_objects`
- `get_binary_content`
- `search_pages`
- `find_meta`
- `get_hyperlink`
- `get_parent`

Hierarchy/list tools hide OneNote recycle-bin items by default. Pass
`include_recycle_bin=true` when you need to inspect deleted pages or sections.

Identifiers accepted by tools are resolved in this order:

1. exact OneNote object ID
2. exact hierarchy path, for example `Notebook/Section Group/Section`
3. unique display name

For automation, prefer IDs or exact paths from `list_hierarchy`,
`list_sections`, or `list_pages`. Display names can be ambiguous.
Call `resolve_identifier` before write or delete operations when you want to
verify that an ID, path, or name resolves to exactly one live object.

`health_check` returns the Python executable, module path, process directory,
identifier resolution order, and default search backend. Use it first when
debugging MCP startup or duplicate checkout issues.

`health_check` also reports supported content formats. The current write
formats are `plain`, `html`, and `markdown`.

Create and update:

- `open_hierarchy`
- `create_notebook`
- `create_section`
- `create_section_group`
- `create_page`
- `update_page_title`
- `append_to_page`
- `add_image_to_page`
- `replace_page_body`
- `delete_page_content`
- `delete_hierarchy`
- `update_page_xml`
- `update_hierarchy_xml`

File/export/app control:

- `publish_object`
- `navigate_to`
- `navigate_to_url`
- `sync_hierarchy`
- `close_notebook`
- `merge_sections`
- `set_filing_location`

`search_pages` uses a live local text scan when `include_unindexed=true`, which
is the default. This finds freshly-created pages without waiting for OneNote's
desktop search index. Pass `include_unindexed=false` when you specifically want
to use the OneNote index.

`get_page_objects` marks objects with `delete_supported` and, for child objects
such as table cells or paragraph OEs, returns `delete_object_id` when the parent
outline is the deletable OneNote COM object. Use that parent ID with
`delete_page_content`.

## Smoke Test

Run a read-only MCP startup and discovery check:

```powershell
.\.venv\Scripts\python.exe scripts\smoke_mcp.py
```

Run a write/read/search smoke test against a chosen section:

```powershell
.\.venv\Scripts\python.exe scripts\smoke_mcp.py --notebook "Notebook" --section "Notebook/Section" --export-dir tmp
```

The script talks to the server through stdio MCP, so it verifies the same path
used by MCP clients instead of importing server functions directly.

## Examples

List all open notebooks:

```text
Use local-onenote health_check, then list_notebooks.
```

Create a page:

```text
Create a page in section "Notebook/Projects/Report" titled "Today's Notes"
with this plain-text body: ...
```

Export a page to PDF:

```text
Publish page "Notebook/Projects/Report/Today's Notes" to
"C:\path\to\exports\today.pdf" as pdf.
```

Use absolute export paths in automation. `publish_object` normalizes relative
paths against the server process directory, but absolute paths make the output
location unambiguous.

Replace a page body:

```text
Use replace_page_body.
```

Add an image:

```text
Use add_image_to_page with image_path. Width and height are optional.
If only one dimension is provided, the server infers the other dimension from
the image's native aspect ratio for PNG, JPEG, GIF, and BMP files.
```

Import Markdown notes:

```text
Use create_page or replace_page_body with content_format="markdown".
The server converts Markdown to HTML through OneMore's Markdig parser, then
emits native OneNote tables where possible.
```

Example Markdown body:

```markdown
# 入学待办

Unique token: example

- 上传照片
- 确认住宿

| 事项 | 状态 |
| --- | --- |
| 体检预约 | **进行中** |
| 学费支付 | 待确认 |
```

## Limits

This server is intentionally local and depends on the OneNote desktop COM API.
It can be stronger than Graph for offline/local notebooks, but Graph still has a
more formal cloud permission model and better multi-user sync semantics.

The server does not directly edit `.one` binary files. That is deliberate:
OneNote's COM API is the stable local write surface.
