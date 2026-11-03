"""Entry point: register every tool and serve MCP over stdio."""

from __future__ import annotations

from . import tools  # noqa: F401  (importing registers the tools)
from .app import mcp


def main() -> None:
    mcp.run(transport="stdio")


if __name__ == "__main__":
    main()
