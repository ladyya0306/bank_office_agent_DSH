#!/usr/bin/env python
"""toolV2 JSON workflow entrypoint: one JSON input and one JSON output."""
from __future__ import annotations
import json
import sys
from pathlib import Path

sys.stdin.reconfigure(encoding="utf-8")
sys.stdout.reconfigure(encoding="utf-8")
sys.stderr.reconfigure(encoding="utf-8")

HERE = Path(__file__).resolve().parent
if str(HERE) not in sys.path: sys.path.insert(0, str(HERE))
from workflow.runner import WorkflowError, dispatch  # noqa: E402

try:
    request = json.load(sys.stdin)
    response = dispatch(request)
except (WorkflowError, ValueError, OSError) as exc:
    response = {"ok": False, "status": "failed", "questions": [], "results": [], "error": str(exc)}
except Exception as exc:
    response = {"ok": False, "status": "failed", "questions": [], "results": [], "error": "未预期流程错误：%s" % exc}
print(json.dumps(response, ensure_ascii=False))
