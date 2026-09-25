"""office_kit — a local office-capability toolkit for the DeepSeek Harness agent.

Every subcommand prints exactly one JSON envelope on stdout:

    {"ok": true, "command": "...", "data": {...}, "artifacts": [{"path": ..., "description": ...}]}

so the agent can parse results and hand produced files to the `present` tool.
Human-readable reports are written as files rather than dumped to stdout, which
keeps tool output small and machine-readable.
"""
from __future__ import annotations

__version__ = "1.0.0"
