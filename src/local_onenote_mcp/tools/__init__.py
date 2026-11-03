"""MCP tools, grouped by what they touch. Importing this package registers them all."""

from . import organizing, reading, system, writing

__all__ = ["organizing", "reading", "system", "writing"]
