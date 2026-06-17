# Local OneNote MCP

A local Microsoft OneNote MCP server for Windows. It controls your OneNote desktop app directly through the local Windows COM API—**no Azure, Microsoft Graph, API keys, or internet-based OAuth required.** Everything stays 100% on your local machine.

## Key Features

- **Local & Offline-First:** Operates directly on your locally installed OneNote desktop application.
- **Rich Interaction:** Read pages, search notebook content, create new sections and pages, and append information.
- **Exporting:** Convert and save OneNote pages directly to PDF.
- **Multimedia Support:** Add local images directly to your OneNote pages.
- **Markdown Support:** Built-in support for converting standard Markdown into native OneNote tables and formatting.

## Requirements

- **Windows**
- **Microsoft OneNote** (Desktop app version, not the Windows 10 app)
- **Python 3.11+**
- **Node.js / npm** (For installation)

## Quick Start

### 1. Install the Server

Open PowerShell and install the package globally:

```powershell
npm install -g github:Peteroooooooo/local-onenote-mcp
```

*(Once published to the npm registry, you will be able to install it via `npm install -g local-onenote-mcp`)*

### 2. Configure Claude Desktop

Add the server to your Claude Desktop configuration file (located at `%APPDATA%\Claude\claude_desktop_config.json`):

```json
{
  "mcpServers": {
    "local-onenote": {
      "command": "local-onenote-mcp",
      "env": {
        "LOCAL_ONENOTE_MCP_TIMEOUT": "90",
        "LOCAL_ONENOTE_MCP_MAX_TEXT_CHARS": "60000"
      }
    }
  }
}
```

### 3. Restart your Client

Restart Claude Desktop. The first time the server starts, it will automatically set up a local Python environment and prepare itself. 

---

## Example Prompts (How to Use)

Once configured, you can ask Claude to perform actions in OneNote using natural language:

### List & Search
- *"List all of my OneNote notebooks."*
- *"Search my OneNote pages for 'Weekly Report'."*
- *"What sections do I have in my 'Projects' notebook?"*

### Create & Update Pages
- *"Create a page titled 'Meeting Notes' in the 'General' section."*
- *"Append this paragraph to my 'To-Do List' page: [insert text]."*
- *"Create a table in a new page in my 'Work' notebook with columns for Task, Status, and Date."*

### Images & Export
- *"Add the image at 'C:\Users\User\Desktop\screenshot.png' to my current page."*
- *"Export my 'Project Specification' page to a PDF file at 'C:\Users\User\Documents\Spec.pdf'."*

---

## Markdown Formatting Example

The server supports standard Markdown formatting (including tables). You can ask Claude to create a page with a body like this:

```markdown
# Weekly Team Sync

- **Date:** 2026-06-17
- **Attendees:** Alice, Bob, Peter

## Key Takeaways
- Q2 targets have been met successfully.
- Code freeze is scheduled for next Thursday.

## Action Items
| Task | Owner | Status |
| --- | --- | --- |
| Update API documentation | Bob | In Progress |
| Prepare release notes | Alice | Pending |
```

---

## Settings & Environment Variables

If you need custom configurations, you can pass these environment variables in your configuration JSON:

- `LOCAL_ONENOTE_MCP_PYTHON`: Path to a specific Python executable if Python is not in your system `PATH` (e.g., `C:\\Python312\\python.exe`).
- `LOCAL_ONENOTE_MCP_TIMEOUT`: Command timeout in seconds (Default: `90`).
- `LOCAL_ONENOTE_MCP_MAX_TEXT_CHARS`: Maximum characters to read from a single page to prevent context overflow (Default: `60000`).
- `LOCAL_ONENOTE_MARKDIG_DLL`: Path to custom `Markdig.Signed.dll` if you use the [OneMore](https://github.com/stevencohn/OneMore) add-in for Markdown-to-HTML parsing. (The server automatically detects standard OneMore installations).
