"""Call the MCP tools of this checkout from the command line, in one server session.

Each run starts `python -m local_onenote_mcp.server` over stdio (so it always runs the
code on disk, not an already running server), calls the tools listed in a JSON file in
order, and prints one JSON result per call. Images returned by a tool (view_page) are
written to files and their paths printed instead of base64.

    .venv\\Scripts\\python.exe scripts\\mcp_session.py calls.json [--images DIR] [--backups FILE]

calls.json holds a list of {"tool": "get_page", "args": {...}}. A call whose args contain
"$N.key.sub" strings gets the value from result N (0-based) of the same run; "${N.key}" inside a
longer string is replaced by that value as text.
"""

from __future__ import annotations

import argparse
import asyncio
import base64
import json
import os
import re
import sys
import tempfile
import time
from pathlib import Path
from typing import Any

from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client

REFERENCE = re.compile(r"^\$(\d+)((?:\.[^.]+)+)$")
INLINE_REFERENCE = re.compile(r"\$\{(\d+)((?:\.[^.}]+)+)\}")


def _lookup(results: list[Any], index: int, path: str) -> Any:
    value: Any = results[index]
    for key in path.strip(".").split("."):
        value = value[int(key)] if isinstance(value, list) else value[key]
    return value


def _resolve(value: Any, results: list[Any]) -> Any:
    if isinstance(value, str):
        match = REFERENCE.match(value)
        if match:
            return _lookup(results, int(match.group(1)), match.group(2))
        return INLINE_REFERENCE.sub(lambda m: str(_lookup(results, int(m.group(1)), m.group(2))), value)
    if isinstance(value, list):
        return [_resolve(item, results) for item in value]
    if isinstance(value, dict):
        return {key: _resolve(item, results) for key, item in value.items()}
    return value


async def run(calls: list[dict[str, Any]], images: Path, backups: str | None) -> list[Any]:
    env = {**os.environ, "LOCAL_ONENOTE_MCP_BACKGROUND_INDEX": "0", "PYTHONIOENCODING": "utf-8"}
    if backups:
        env["LOCAL_ONENOTE_MCP_BACKUP_PATH"] = backups
    params = StdioServerParameters(command=sys.executable, args=["-m", "local_onenote_mcp.server"], env=env)
    results: list[Any] = []
    async with stdio_client(params) as (read, write):
        async with ClientSession(read, write) as session:
            await session.initialize()
            for number, call in enumerate(calls):
                started = time.perf_counter()
                try:
                    args = _resolve(call.get("args", {}), results)
                    result = await session.call_tool(call["tool"], args)
                except Exception as exc:  # a bad reference or a transport error
                    results.append({"ok": False, "error": f"{type(exc).__name__}: {exc}"})
                    continue
                elapsed = round((time.perf_counter() - started) * 1000)
                texts, saved = [], []
                for block in result.content:
                    if block.type == "image":
                        target = images / f"call{number}-{len(saved)}.{block.mimeType.split('/')[-1]}"
                        target.write_bytes(base64.b64decode(block.data))
                        saved.append(str(target))
                    elif block.type == "text":
                        texts.append(block.text)
                if result.isError:
                    entry: Any = {"ok": False, "error": "\n".join(texts)}
                elif saved:
                    entry = {"ok": True, "text": "\n".join(texts), "images": saved}
                else:
                    try:
                        entry = json.loads(texts[0]) if texts else {}
                    except json.JSONDecodeError:
                        entry = {"ok": True, "text": "\n".join(texts)}
                if isinstance(entry, dict):
                    entry["_ms"] = elapsed
                results.append(entry)
    return results


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("calls", help="JSON file with a list of {tool, args}")
    parser.add_argument("--images", default="", help="folder for images returned by tools")
    parser.add_argument("--backups", default="", help="backup database to use instead of the default one")
    options = parser.parse_args()
    calls = json.loads(Path(options.calls).read_text(encoding="utf-8"))
    images = Path(options.images or tempfile.mkdtemp(prefix="onenote-mcp-images-"))
    images.mkdir(parents=True, exist_ok=True)
    results = asyncio.run(run(calls, images, options.backups or None))
    sys.stdout.reconfigure(encoding="utf-8")
    for number, (call, result) in enumerate(zip(calls, results)):
        print(json.dumps({"call": number, "tool": call["tool"], "result": result}, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
