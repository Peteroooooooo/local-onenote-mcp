"""Run a small end-to-end smoke test through the MCP stdio transport.

By default this only verifies startup and read-only discovery. Pass --section
to create one test page in that section and verify write/read/search behavior.
"""

from __future__ import annotations

import argparse
import asyncio
import json
import sys
from datetime import datetime
from pathlib import Path
from typing import Any

from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Smoke test the local OneNote MCP server.")
    parser.add_argument(
        "--server-python",
        default=sys.executable,
        help="Python executable used to run -m local_onenote_mcp.server.",
    )
    parser.add_argument(
        "--notebook",
        default="",
        help="Optional notebook identifier to list. Use an ID, exact path, or unique name.",
    )
    parser.add_argument(
        "--section",
        default="",
        help="Section identifier for write testing. Use an ID, exact path, or unique name.",
    )
    parser.add_argument(
        "--export-dir",
        default="",
        help="Optional directory for exporting the smoke test page as PDF.",
    )
    parser.add_argument(
        "--delete-page",
        action="store_true",
        help="Move the created smoke test page to OneNote recycle bin after verification.",
    )
    return parser.parse_args()


def text_of(result: Any) -> str:
    return "\n".join(getattr(content, "text", str(content)) for content in result.content)


def parse_tool_result(result: Any) -> dict[str, Any]:
    text = text_of(result)
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        return {"ok": False, "error": f"Tool returned non-JSON text: {text[:500]}"}


async def call_tool(session: ClientSession, name: str, args: dict[str, Any]) -> dict[str, Any]:
    return parse_tool_result(await session.call_tool(name, args))


async def run_smoke(args: argparse.Namespace) -> dict[str, Any]:
    params = StdioServerParameters(
        command=args.server_python,
        args=["-m", "local_onenote_mcp.server"],
        env={
            "LOCAL_ONENOTE_MCP_TIMEOUT": "90",
            "LOCAL_ONENOTE_MCP_MAX_TEXT_CHARS": "60000",
        },
    )
    checks: list[dict[str, Any]] = []
    failures: list[str] = []

    async with stdio_client(params) as (read, write):
        async with ClientSession(read, write) as session:
            await session.initialize()

            tools = await session.list_tools()
            tool_names = {tool.name for tool in tools.tools}
            required_tools = {
                "health_check",
                "resolve_identifier",
                "list_notebooks",
                "list_sections",
                "create_page",
                "search_pages",
            }
            missing_tools = sorted(required_tools - tool_names)
            checks.append({"name": "list_tools", "ok": not missing_tools, "tool_count": len(tool_names)})
            if missing_tools:
                failures.append(f"Missing required tools: {', '.join(missing_tools)}")

            health = await call_tool(session, "health_check", {})
            checks.append({"name": "health_check", "ok": health.get("ok"), "result": health})
            if not health.get("ok"):
                failures.append(f"health_check failed: {health.get('error')}")

            notebooks = await call_tool(session, "list_notebooks", {})
            notebook_found = True
            if args.notebook:
                notebook_found = any(
                    item.get("name") == args.notebook or item.get("id") == args.notebook
                    for item in notebooks.get("notebooks", [])
                )
            checks.append({"name": "list_notebooks", "ok": notebooks.get("ok") and notebook_found})
            if not notebooks.get("ok"):
                failures.append(f"list_notebooks failed: {notebooks.get('error')}")
            elif args.notebook and not notebook_found:
                failures.append(f"Notebook not found by name or ID: {args.notebook}")

            sections = await call_tool(session, "list_sections", {"notebook_identifier": args.notebook} if args.notebook else {})
            checks.append({"name": "list_sections", "ok": sections.get("ok"), "count": sections.get("count")})
            if not sections.get("ok"):
                failures.append(f"list_sections failed: {sections.get('error')}")

            if args.notebook:
                resolved_notebook = await call_tool(
                    session,
                    "resolve_identifier",
                    {"identifier": args.notebook, "item_type": "notebook"},
                )
                checks.append({"name": "resolve_identifier:notebook", "ok": resolved_notebook.get("ok")})
                if not resolved_notebook.get("ok"):
                    failures.append(f"resolve_identifier notebook failed: {resolved_notebook.get('error')}")

            if not args.section:
                return {"ok": not failures, "mode": "read_only", "checks": checks, "failures": failures}

            pages = await call_tool(session, "list_pages", {"section_identifier": args.section})
            checks.append({"name": "list_pages", "ok": pages.get("ok"), "count": pages.get("count")})
            if not pages.get("ok"):
                failures.append(f"list_pages failed: {pages.get('error')}")
                return {"ok": False, "mode": "write", "checks": checks, "failures": failures}

            resolved_section = await call_tool(
                session,
                "resolve_identifier",
                {"identifier": args.section, "item_type": "section"},
            )
            checks.append({"name": "resolve_identifier:section", "ok": resolved_section.get("ok")})
            if not resolved_section.get("ok"):
                failures.append(f"resolve_identifier section failed: {resolved_section.get('error')}")
                return {"ok": False, "mode": "write", "checks": checks, "failures": failures}

            stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
            title = f"OneNote MCP smoke {stamp}"
            token = f"MCP_SMOKE_TOKEN_{stamp}"
            content = (
                f"<p>Smoke token: {token}</p>"
                "<table>"
                "<tr><th>Feature</th><th>Status</th></tr>"
                "<tr><td>Create page</td><td>pass</td></tr>"
                "<tr><td>Native table</td><td>pass</td></tr>"
                "</table>"
            )
            created = await call_tool(
                session,
                "create_page",
                {"section_identifier": args.section, "title": title, "content": content, "content_format": "html"},
            )
            page_id = created.get("page_id")
            checks.append({"name": "create_page", "ok": created.get("ok"), "page_id": page_id})
            if not created.get("ok"):
                failures.append(f"create_page failed: {created.get('error')}")
                return {"ok": False, "mode": "write", "checks": checks, "failures": failures}

            page = await call_tool(session, "get_page", {"page_identifier": page_id, "include_xml": True, "page_info": "all"})
            page_xml = page.get("xml", "")
            page_ok = page.get("ok") and token in page.get("text", "") and "<one:Table" in page_xml
            checks.append({"name": "get_page", "ok": page_ok})
            if not page_ok:
                failures.append("get_page did not return the smoke token and native one:Table XML.")

            search = await call_tool(
                session,
                "search_pages",
                {
                    "query": token,
                    "start_identifier": args.notebook,
                    "max_results": 5,
                    "include_unindexed": True,
                },
            )
            found = any(item.get("id") == page_id for item in search.get("pages", []))
            checks.append({"name": "search_pages", "ok": search.get("ok") and found, "backend": search.get("search_backend")})
            if not found:
                failures.append(f"search_pages did not find the created page for token {token}.")

            if args.export_dir:
                export_path = Path(args.export_dir).expanduser().resolve(strict=False) / f"onenote-mcp-smoke-{stamp}.pdf"
                exported = await call_tool(
                    session,
                    "publish_object",
                    {"object_identifier": page_id, "target_path": str(export_path), "format": "pdf", "overwrite": True},
                )
                export_ok = exported.get("ok") and export_path.exists() and export_path.stat().st_size > 1000
                checks.append({"name": "publish_object", "ok": export_ok, "path": str(export_path)})
                if not export_ok:
                    failures.append(f"publish_object did not create a usable PDF at {export_path}.")

            if args.delete_page:
                deleted = await call_tool(session, "delete_hierarchy", {"object_identifier": page_id, "permanently": False})
                checks.append({"name": "delete_hierarchy", "ok": deleted.get("ok")})
                if not deleted.get("ok"):
                    failures.append(f"delete_hierarchy failed: {deleted.get('error')}")

    return {"ok": not failures, "mode": "write" if args.section else "read_only", "checks": checks, "failures": failures}


def main() -> int:
    result = asyncio.run(run_smoke(parse_args()))
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0 if result["ok"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
