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
   .\.venv\Scripts\python.exe scripts\e2e_onenote.py          # full check in throwaway notebooks
   ```

Starting the server does not start OneNote. OneNote is contacted on the first tool call.

---

## Markdown, HTML & Reading Formats

Write tools (`create_page`, `append_to_page`, `edit_page`, `replace_page_body`) accept `content_format` = `plain`, `html`, or `markdown`. Markdown is parsed with the bundled `markdown-it-py` (CommonMark plus tables and strikethrough). It becomes:

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

`get_page` takes `format` = `text` or `markdown` and `include_ocr` (default `true`). Markdown read-back maps the same structures back (headings, lists, `- [x]` to-dos, tables, code fences, image OCR text as quotes). Handwriting that OneNote has recognised is read as text; other ink is marked `[Handwriting: …]`. To see such a page, `view_page` prints it to PDF in OneNote and returns each printed page as a JPEG image (about 1.3 s for a one-page note, 240 KB per image at the default 1000 px width).

---

## Editing part of a page

`edit_page` changes single paragraphs and leaves everything else on the page alone: images, handwriting, attachments and other text blocks are not resubmitted. Take paragraph IDs from `get_page(include_paragraphs=true)`, then send a list of operations that are applied together in one update:

```json
[
  {"op": "replace_text", "find": "2025", "replace": "2026"},
  {"op": "set_text", "paragraph_id": "{…}{48}{B0}", "content": "**new** wording"},
  {"op": "insert_after", "paragraph_id": "{…}{51}{B0}", "content": "- extra item\n- [ ] follow-up"},
  {"op": "set_todo", "paragraph_id": "{…}{55}{B0}", "done": true},
  {"op": "delete", "paragraph_id": "{…}{60}{B0}"},
  {"op": "add_row", "paragraph_id": "{…}{70}{B0}", "cells": ["Gamma", "QA", "Planned"]}
]
```

- `replace_text` keeps bold, links and other formatting around the match; a match must lie inside one formatting run.
- `set_text` keeps the paragraph's bullet, number, to-do box and style.
- `insert_after` / `insert_before` take any amount of Markdown, HTML or plain text and reuse the page's own heading and To Do definitions.
- `delete` removes the paragraph and its indented children; a block left empty is removed.
- Table cells are paragraphs too: `get_page(include_paragraphs=true)` gives each one its `table`, `row` and `column`. `set_text` changes a cell, `add_row` adds a row after the row of the given cell (or at the end, given the table's own paragraph), `delete_row` removes the row of the given cell. Deleting the only paragraph of a cell empties it.

If one operation is invalid (for example an unknown paragraph ID), nothing is written.

---

## Backups

Every destructive tool (`edit_page`, `replace_page_body`, `delete_page_content`, `update_page_xml`, `delete_hierarchy`, `restore_backup` itself) first saves the page to `%LOCALAPPDATA%\local-onenote-mcp\backups.sqlite3`. `list_backups` shows them and `restore_backup` brings one back, either into the original page (`mode="replace"`) or as a new page (`mode="new_page"`, which also works after the page was deleted). The default, `mode="auto"`, puts block backups back in place and restores whole pages as new pages.

The store stays small and bounded:

- **Images, ink and attachments are stored once.** Each is kept by its SHA-256 hash; ten backups of a page with the same photo store the photo once. The page XML itself is zlib-compressed.
- **`edit_page` saves only the blocks it changes** (backup kind `blocks`): the text block's XML, a few KB, without reading the page's images. If a changed block holds an image, ink or a file, the whole page is saved instead.
- **Unchanged pages are not saved again.** If the page is identical to its newest backup, that backup is reused.
- **Old backups are dropped automatically:** at most 10 per page, none older than 30 days, and the whole file under 200 MB (oldest first; the newest backup is always kept). Images that no backup uses any more are deleted and the file shrinks.

Measured on 150 random pages of a real 4,800-page collection: the page text and structure take 5 KB per backup on average. Images and ink take the rest (about 0.8 MB per page on average, already compressed by their formats). Saving takes 44 ms (median), and saving unchanged pages again adds 0 bytes. Change the limits with the variables below, or set `LOCAL_ONENOTE_MCP_BACKUP=0` to turn backups off.

---

## Tool Directory

Each tool is annotated as read-only, writing, or destructive, so MCP clients can auto-approve reads and ask before deletes. Failures are returned as MCP tool errors (`isError: true`) with the reason.

### 1. Discovery & Reading
* `health_check`: Server version, Python location, OneNote's special folders, text-index progress and backup storage.
* `resolve_identifier`: Turn an ID, path, or name into one live object.
* `list_notebooks` / `list_sections` / `list_pages` / `list_hierarchy`: Traverse the live OneNote structure. `list_hierarchy` returns at most `max_items` (default 500) and reports `total` / `truncated`.
* `get_page`: Title and body as text or Markdown. `include_paragraphs=true` lists paragraph IDs for `edit_page`; `include_objects=true` lists outlines, images, ink and attachments with `object_id` and `callback_id`; `include_xml=true` adds the raw page XML.
* `get_page_image`: Return an image from a page as MCP image content, so the model can look at it.
* `view_page`: Show the whole page as it looks in OneNote (handwriting, drawings, layout) as JPEG images, one per printed page.
* `get_binary_content`: Download an image (`callback_id`) or attached file (`file_name`), as base64 or straight to a file with `save_to`.
* `search_pages`: Pages containing every word of the query, ranked (title matches first, then number of occurrences in the body). `modified_after` / `modified_before` filter by edit date; with no query they list the pages edited in that window, newest first.
* `get_hyperlink`

### 2. Creation & Edits
* `open_hierarchy` / `create_notebook` / `create_section` / `create_section_group`
* `create_page`, `update_page_title`, `append_to_page`, `add_image_to_page`, `add_file_to_page`
* `edit_page`: Change, insert, delete or tick single paragraphs, and edit table cells and rows (see above).
* `copy_page`: Copy a page with its images, handwriting and attachments into any section.
* `replace_page_body`: Replace the whole body. Writes the new body first, then removes the old objects, so a failed write never empties the page.
* `delete_page_content`, `delete_hierarchy` (pages are always matched by ID, so a page with a duplicate title is never touched)
* `list_backups`, `restore_backup`

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
| `LOCAL_ONENOTE_MCP_BACKUP` | `1` | `0` turns page backups off |
| `LOCAL_ONENOTE_MCP_BACKUP_PATH` | | Custom location for the backup file |
| `LOCAL_ONENOTE_MCP_BACKUP_MB` | `200` | Size limit of the backup file |
| `LOCAL_ONENOTE_MCP_BACKUP_VERSIONS` | `10` | Backups kept per page |
| `LOCAL_ONENOTE_MCP_BACKUP_DAYS` | `30` | Backups older than this are deleted |

---

## Code Layout

| File | Role |
| --- | --- |
| `server.py` | Entry point: registers the tools and serves MCP over stdio |
| `app.py`, `config.py` | Shared runtime (settings, OneNote connection, index, backups) and the `tool` decorator |
| `bridge.py`, `onenote.py` | The PowerShell COM bridge; hierarchy cache, identifier resolution and page I/O on top of it |
| `hierarchy.py`, `pagexml.py`, `markup.py`, `reader.py` | Parse hierarchy XML; build and inspect page XML; Markdown/HTML to OneNote; OneNote to text/Markdown |
| `editing.py` | Paragraph-level page edits |
| `backup.py`, `pages.py` | Backup store; snapshots, copies and restores |
| `search.py`, `text_index.py` | Search ranking, date filters and the background text index |
| `tools/` | The MCP tools: `reading`, `writing`, `organizing`, `system` |

---

## Tests

```powershell
# Unit tests (an in-memory fake stands in for OneNote; one bridge test starts PowerShell)
.\.venv\Scripts\python.exe -m pytest -q

# Full cycle in throwaway notebooks (temp folder, closed afterwards)
.\.venv\Scripts\python.exe scripts\e2e_onenote.py
```

`e2e_onenote.py` runs 30 checks against the real OneNote app, including partial edits of text and tables, block and whole-page restores, `view_page`, copying a page with an image and an attachment, moving pages and sections between two notebooks, and confirming the result after closing and reopening them. It also exports its test page to PDF and prints the folder, so formatting can be checked by eye.

---

## Limits & Boundaries

OneNote sometimes answers "class not registered" for a moment while it rewrites notebook files; reads and hierarchy updates are retried once after 0.5 s.

This server relies on the Windows COM API, so it runs only on Windows, for the signed-in local user, with the OneNote desktop app installed. COM calls are serialized through one bridge process; a tool call waits for at most one background-index page read. OneNote's own search index can lag behind very recent edits, which is why `search_pages` falls back to a text scan.
