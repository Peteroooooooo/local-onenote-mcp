"""Validate the Codex local-onenote MCP configuration points at this checkout."""

from __future__ import annotations

import json
import subprocess
import sys
import tomllib
from pathlib import Path
from typing import Any


SERVER_NAME = "local-onenote"


def load_config(path: Path) -> dict[str, Any]:
    with path.open("rb") as handle:
        return tomllib.load(handle)


def main() -> int:
    repo_root = Path(__file__).resolve().parents[1]
    config_path = Path.home() / ".codex" / "config.toml"
    checks: list[dict[str, Any]] = []
    failures: list[str] = []

    if not config_path.is_file():
        result = {"ok": False, "config_path": str(config_path), "failures": ["Codex config.toml was not found."]}
        print(json.dumps(result, ensure_ascii=False, indent=2))
        return 1

    config = load_config(config_path)
    server = config.get("mcp_servers", {}).get(SERVER_NAME)
    checks.append({"name": "server_entry", "ok": isinstance(server, dict)})
    if not isinstance(server, dict):
        failures.append(f"Missing [mcp_servers.{SERVER_NAME}] in {config_path}.")
        result = {"ok": False, "config_path": str(config_path), "checks": checks, "failures": failures}
        print(json.dumps(result, ensure_ascii=False, indent=2))
        return 1

    command = Path(str(server.get("command", ""))).expanduser()
    args = server.get("args", [])
    command_exists = command.is_file()
    args_ok = args == ["-m", "local_onenote_mcp.server"]
    checks.append({"name": "command_exists", "ok": command_exists, "command": str(command)})
    checks.append({"name": "args", "ok": args_ok, "args": args})
    if not command_exists:
        failures.append(f"Configured command does not exist: {command}")
    if not args_ok:
        failures.append('Expected args to be ["-m", "local_onenote_mcp.server"].')

    module_path = ""
    module_under_repo = False
    if command_exists:
        probe = (
            "from pathlib import Path; "
            "import local_onenote_mcp.server as s; "
            "print(Path(s.__file__).resolve())"
        )
        completed = subprocess.run(
            [str(command), "-c", probe],
            text=True,
            capture_output=True,
            timeout=30,
        )
        if completed.returncode == 0:
            module_path = completed.stdout.strip()
            try:
                Path(module_path).relative_to(repo_root)
                module_under_repo = True
            except ValueError:
                module_under_repo = False
        else:
            failures.append((completed.stderr or completed.stdout or "Import probe failed.").strip())
        checks.append({"name": "module_import", "ok": completed.returncode == 0, "module_path": module_path})
        checks.append({"name": "module_under_this_repo", "ok": module_under_repo, "repo_root": str(repo_root)})
        if completed.returncode == 0 and not module_under_repo:
            failures.append(f"Configured Python imports local_onenote_mcp from another checkout: {module_path}")

    result = {
        "ok": not failures,
        "config_path": str(config_path),
        "server": SERVER_NAME,
        "checks": checks,
        "failures": failures,
    }
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0 if result["ok"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
