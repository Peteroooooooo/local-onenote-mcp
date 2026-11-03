# Local OneNote MCP

A local Microsoft OneNote MCP server for Windows. It controls the OneNote desktop app directly through the local OneNote COM API—**requiring no Azure, Microsoft Graph, API keys, or online OAuth.**

---

## Design & Architecture

- **Local-Only Boundary:** Every operation executes directly through the local OneNote desktop installation. No data leaves your computer.
- **COM-First Engineering:** No direct binary `.one` file manipulation. All writes and reads go through OneNote's own COM engine, so sync and integrity stay OneNote's job.
- **Persistent, Injection-Free Bridge:** One long-lived PowerShell process runs a fixed script and receives each request as a JSON line. User data is never interpolated into script text. Keeping the process (and its COM connection) alive makes a typical page read take tens of milliseconds instead of about a second.
- **Native Formatting:** Markdown and HTML become native OneNote headings, bullet/numbered lists (nested), To Do checkboxes, tables and code blocks. Pages read back as plain text or Markdown, including the text OneNote recognised inside images (OCR).
- **Local Full-Text Index:** Page text (with OCR) is cached per page and refreshed when OneNote reports the page changed, so searching thousands of pages takes well under a second once the index is built.

> **Design Note:** PowerShell is used as the COM bridge because Python automation libraries (pywin32) cannot build a dispatch wrapper for `OneNote.Application`, while PowerShell can call it directly.

---

## Requirements

- **Windows 10 / 11**
- **Microsoft OneNote Desktop App** (Traditional version; not the legacy Windows 10 UWP app)
- **Python 3.11+**

---

## Setup (local folder)

1. **Clone the repository and create its virtual environment:**
   ```powershell
   git clone https://github.com/Peteroooooooo/local-onenote-mcp
   cd local-onenote-mcp
   python -m venv .venv
   .\.venv\Scripts\python.exe -m pip install -e .
   ```
   The editable install means a `git pull` or local edit takes effect the next time the MCP client starts the server.

2. **Point your MCP client at the folder's Python.** Replace the path with your clone.

   Claude Code / Claude Desktop (JSON):
   ```json
   {
     "mcpServers": {
       "local-onenote": {
         "type": "stdio",
         "command": "C:\\path\\to\\local-onenote-mcp\\.venv\\Scripts\\python.exe",
         "args": ["-m", "local_onenote_mcp.server"],
         "env": {
           "LOCAL_ONENOTE_MCP_TIMEOUT": "90",
           "LOCAL_ONENOTE_MCP_MAX_TEXT_CHARS": "60000"
         }
       }
     }
   }
   ```

   Codex (TOML; use forward slashes or doubled backslashes):
   ```toml
   [mcp_servers.local-onenote]
   type = "stdio"
   command = "C:/path/to/local-onenote-mcp/.venv/Scripts/python.exe"
   args = ["-m", "local_onenote_mcp.server"]
   startup_timeout_ms = 30000
   ```

3. **Check it:**
   ```powershell
   .\.venv\Scripts\python.exe scripts\check_codex_config.py   # Codex config points at this folder
   .\.venv\Scripts\python.exe scripts\smoke_mcp.py            # read-only check against your OneNote
   ```

Starting the server does not start OneNote. OneNote is contacted on the first tool call.

> The repository also contains an npm launcher (`bin/local-onenote-mcp.js`) that builds its own cached virtual environment. It is optional; the local-folder setup above is the supported path.

---

## Markdown, HTML & Reading Formats

Write tools (`create_page`, `append_to_page`, `replace_page_body`) accept `content_format` = `plain`, `html`, or `markdown`. Markdown is parsed with the bundled `markdown-it-py` (CommonMark plus tables and strikethrough). It becomes:

| Markdown / HTML | In OneNote |
| --- | --- |
| `#` … `######` headings | OneNote Heading 1–6 styles |
| `-` / `1.` lists, nested lists | Native bullet / numbered lists with indentation |
| `- [ ]` / `- [x]` | Native To Do tags you can tick in OneNote |
| `>` quotes | OneNote quote style |
| Tables | Native OneNote tables with a shaded header row |
| Fenced code | OneNote code style, one line per paragraph, indentation kept |
| `**bold**`, `*italic*`, `~~strike~~`, `` `code` ``, links | Inline formatting (`http`, `https`, `mailto`, `onenote` links only) |

`append_to_page` and `add_image_to_page` place the new block below the page's existing content unless you pass `x`/`y`.

`get_page` takes `format` = `text` or `markdown` and `include_ocr` (default `true`). Markdown read-back maps the same structures back (headings, lists, `- [x]` to-dos, tables, code fences, image OCR text as quotes).

---

## Tool Directory

Each tool is annotated as read-only, writing, or destructive, so MCP clients can auto-approve reads and ask before deletes. Failures are returned as MCP tool errors (`isError: true`) with the reason.

### 1. Discovery & Reading
* `health_check`: Server version, Python location, feature list, and text-index progress.
* `resolve_identifier`: Turn an ID, path, or name into one live object.
* `list_notebooks` / `list_sections` / `list_pages` / `list_hierarchy`: Traverse the live OneNote structure. `list_hierarchy` returns at most `max_items` (default 500) and reports `total` / `truncated`.
* `get_page`: Title and body as text or Markdown. `include_objects=true` lists outlines, images, ink and attachments with `object_id` and `callback_id`; `include_xml=true` adds the raw page XML.
* `get_page_image`: Return an image from a page as MCP image content, so the model can look at it.
* `get_binary_content`: Download an image or attached file as base64.
* `search_pages`: Pages containing every word of the query, ranked (title matches first, then number of occurrences in the body).
* `find_meta`, `get_hyperlink`, `get_special_locations`

### 2. Creation & Edits
* `open_hierarchy` / `create_notebook` / `create_section` / `create_section_group`
* `create_page`, `update_page_title`, `append_to_page`, `add_image_to_page`
* `replace_page_body`: Writes the new body first, then removes the old objects, so a failed write never empties the page.
* `delete_page_content`, `delete_hierarchy` (pages are always matched by ID, so a page with a duplicate title is never touched)

### 3. Moving & Organizing
* `move_page`: Move a page to another section in any notebook, or reorder it inside its section (`after_page_identifier`, `to_top`, or end). Subpages move with it unless `include_subpages=false`. A page moved to another section gets a new ID; the tool returns old and new IDs.
* `set_page_level`: Make a page a subpage (level 2 or 3) or promote it back to level 1.
* `move_section`: Move a section or section group into a notebook or section group in any notebook, or reorder it among its siblings. IDs are kept. For local notebooks it waits until OneNote has moved the `.one` file on disk, so the move survives closing the notebook.
* `rename_section`: Rename a section or section group (pages are renamed with `update_page_title`).
* `merge_sections`: Merge one section into another.

All of these refuse to touch the recycle bin, and the section tools refuse name clashes and moving a group into itself.

### 4. File & App Control
* `publish_object`: Export a notebook, section, or page to PDF, XPS, Word, MHT, HTML, EMF, or OneNote package files.
* `navigate_to` / `navigate_to_url`, `sync_hierarchy`, `close_notebook`, `set_filing_location`

### 5. Raw Low-Level Control
* `update_page_xml` / `update_hierarchy_xml`: Submit raw OneNote XML to `UpdatePageContent` / `UpdateHierarchy`.

> **Identifier Resolution:** identifiers are tried as
> 1. OneNote object ID (fastest; the only safe choice for pages, because page titles are often duplicated)
> 2. Hierarchy path (e.g., `Personal/Quick Notes/My Section`)
> 3. Unique display name

### How search works

`search_pages` with the default `backend="auto"`:

1. If the local text index covers the pages in scope, it answers from the index: complete, includes image OCR, about 50 ms for 4,800 pages. Pages edited since they were indexed are re-read on the spot.
2. Otherwise it asks OneNote's own search index, and falls back to scanning page text (time-limited, reports `scan_complete`) when that finds nothing.

The index fills itself in the background after the first tool call and keeps up by re-reading only pages whose `lastModifiedTime` changed. Progress is shown in `health_check` → `text_index`. `backend="local_scan"` forces the scan; `backend="index"` uses only OneNote's index.

The index is stored at `%LOCALAPPDATA%\local-onenote-mcp\text-index.sqlite3` and contains the plain text of your pages. Delete the file to reset it, or set `LOCAL_ONENOTE_MCP_INDEX=0` to keep it in memory only.

### Environment Variables

| Variable | Default | Meaning |
| --- | --- | --- |
| `LOCAL_ONENOTE_MCP_TIMEOUT` | `90` | Seconds before a single COM call is abandoned and the bridge restarted |
| `LOCAL_ONENOTE_MCP_MAX_TEXT_CHARS` | `60000` | Default cap on returned page text |
| `LOCAL_ONENOTE_MCP_SCAN_SECONDS` | `60` | Time budget for a text-scan search |
| `LOCAL_ONENOTE_MCP_HIERARCHY_CACHE_SECONDS` | `5` | How long a hierarchy listing is reused between tool calls (cleared after any write; `0` disables) |
| `LOCAL_ONENOTE_MCP_INDEX` | `1` | `0` keeps the text index in memory only (nothing written to disk) |
| `LOCAL_ONENOTE_MCP_INDEX_PATH` | | Custom location for the text index file |
| `LOCAL_ONENOTE_MCP_BACKGROUND_INDEX` | `1` | `0` stops the background indexer; pages are then indexed only when a search reads them |
| `LOCAL_ONENOTE_MCP_INDEX_REFRESH_SECONDS` | `120` | Minimum gap between background passes that pick up edited pages |

---

## Tests

```powershell
# Unit tests (no OneNote needed except one bridge test that only starts PowerShell)
.\.venv\Scripts\python.exe -m pytest -q

# Read-only check against your OneNote
.\.venv\Scripts\python.exe scripts\smoke_mcp.py

# Full write/read/search/image/move/delete cycle in throwaway notebooks (temp folder, closed afterwards)
.\.venv\Scripts\python.exe scripts\e2e_onenote.py
```

`e2e_onenote.py` runs 20 checks, including moving pages and sections between two notebooks and confirming the result after closing and reopening them. It also exports its test page to PDF and prints the folder, so formatting can be checked by eye.

---

## Limits & Boundaries

This server relies on the Windows COM API, so it runs only on Windows, for the signed-in local user, with the OneNote desktop app installed. COM calls are serialized through one bridge process; a tool call waits for at most one background-index page read. OneNote's own search index can lag behind very recent edits, which is why `search_pages` falls back to a text scan.
