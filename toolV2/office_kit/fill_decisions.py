"""Persist fill-card answers in the same database used by db-fill.

The row fingerprint deliberately excludes run IDs and row numbers. It includes
the batch, exact template bytes, target, proposed value and its provenance, so
a new conversation can reuse a decision while changed evidence cannot.
"""
from __future__ import annotations

import hashlib
import json
import sys
from pathlib import Path

from .common import OfficeKitError
from .harness import _guard_fill, parse_fill_selection
from .store_v2 import StoreV2, now_utc, write_event


def needs_answer(row: dict) -> bool:
    return row.get("kind") == "slot" and (
        row.get("decision") in ("ask", "auto") or
        (row.get("decision") == "empty" and bool(row.get("is_required")))
    )


def fingerprint(plan: dict, row: dict) -> str:
    template_hash = row.get("template_sha256")
    if not template_hash:
        raise OfficeKitError("这份预演清单生成于旧版本，请重新预演后再保存确认。")
    data = {
        "batch": plan.get("batch_no"), "template_sha256": template_hash,
        "field": row.get("field"), "label": row.get("label"),
        "target": row.get("target"), "value": row.get("value"),
        "entity_id": row.get("entity_id"), "entity_name": row.get("entity_name"),
        "subject_eid": row.get("subject_eid"),
        "decision": row.get("decision"), "ask_reason": row.get("ask_reason"),
        "provenance": row.get("provenance"), "source_kind": row.get("source_kind"),
        "is_required": bool(row.get("is_required")),
        "ambiguous": bool(row.get("ambiguous")),
        "candidates": row.get("candidates") if row.get("ambiguous") else None,
    }
    if row.get("subject_scope"):
        data["subject_scope"] = row["subject_scope"]
    raw = json.dumps(data, ensure_ascii=False, sort_keys=True, separators=(",", ":"))
    return hashlib.sha256(raw.encode("utf-8")).hexdigest()


def _rows(plan: dict) -> list[dict]:
    return [row for row in plan.get("rows", []) if needs_answer(row)]


def _store(plan: dict) -> StoreV2:
    db_path = plan.get("db_path")
    if not db_path or not Path(db_path).is_file():
        raise OfficeKitError("预演清单没有可用的办公数据库，请重新预演。")
    return StoreV2(db_path, actor="user", create=False)


def saved_choices(plan: dict) -> tuple[dict, list[int]]:
    """Return CLI arguments for unchanged confirmations and unanswered row IDs."""
    choices: dict = {"apply_all": True, "select": "", "blank": [], "new": [], "use": []}
    selected: list[str] = []
    missing: list[int] = []
    with _store(plan) as store:
        for row in _rows(plan):
            hit = store.conn.execute(
                "SELECT action,value FROM fill_decision WHERE fingerprint=?",
                (fingerprint(plan, row),)).fetchone()
            if hit is None:
                missing.append(int(row["n"]))
                continue
            n = str(row["n"])
            action, value = hit["action"], hit["value"]
            if action == "accept":
                selected.append(n)
            elif action == "blank":
                choices["blank"].append(n)
            elif action == "new" and value is not None:
                choices["new"].append(f"{n}={value}")
            elif action == "use" and value is not None:
                choices["use"].append(f"{n}={value}")
            else:
                missing.append(int(n))
    choices["select"] = ",".join(selected)
    return choices, missing


def save_choices(plan: dict, choices: dict) -> int:
    """Validate every visible question, then save all answers as one transaction."""
    selection = parse_fill_selection(
        choices.get("select"), bool(choices.get("apply_all")),
        choices.get("new"), choices.get("blank"), choices.get("use"))
    rows = _rows(plan)
    unanswered = [row["n"] for row in rows if row["n"] not in selection["actions"]]
    if unanswered:
        raise OfficeKitError("还有 %d 项没有得到使用者回答，未保存。" % len(unanswered))
    _, problems = _guard_fill(plan, selection)
    if problems:
        raise OfficeKitError("确认选项与预演不符：%s" % problems[0]["why"])
    if not rows:
        return 0
    store = _store(plan)
    try:
        store.conn.execute("BEGIN IMMEDIATE")
        for row in rows:
            action = selection["actions"][row["n"]]
            store.conn.execute(
                "INSERT INTO fill_decision(fingerprint,batch_no,template_sha256,field,"
                "action,value,decided_at,run_id) VALUES(?,?,?,?,?,?,?,?) "
                "ON CONFLICT(fingerprint) DO UPDATE SET action=excluded.action,"
                "value=excluded.value,decided_at=excluded.decided_at,run_id=excluded.run_id",
                (fingerprint(plan, row), plan["batch_no"], row["template_sha256"],
                 row["field"], action["action"], action.get("value"),
                 now_utc(), plan["run_id"]))
        write_event(store.conn, batch_no=plan["batch_no"],
                    event_type="fill_review_confirmed", actor_id="user",
                    run_id=plan["run_id"], target="填报确认",
                    payload={"count": len(rows), "answers": [
                        {"field": row["field"], "template": row["template"],
                         "action": selection["actions"][row["n"]]["action"],
                         "value": selection["actions"][row["n"]].get("value")}
                        for row in rows]})
        store.conn.commit()
        return len(rows)
    except Exception:
        store.conn.rollback()
        raise
    finally:
        store.conn.close()


def main() -> int:
    if len(sys.argv) != 3 or sys.argv[1] not in ("lookup", "save"):
        print(json.dumps({"ok": False, "error": "用法：lookup|save fill_plan.json"},
                         ensure_ascii=False))
        return 2
    try:
        plan = json.loads(Path(sys.argv[2]).read_text(encoding="utf-8"))
        if plan.get("stage") != "plan" or not isinstance(plan.get("rows"), list):
            raise OfficeKitError("不是有效的填报预演清单")
        if sys.argv[1] == "lookup":
            choices, missing = saved_choices(plan)
            result = {"ok": True, "arguments": choices, "missing": missing}
        else:
            choices = json.load(sys.stdin)
            result = {"ok": True, "saved": save_choices(plan, choices)}
        print(json.dumps(result, ensure_ascii=False))
        return 0
    except Exception as exc:
        print(json.dumps({"ok": False, "error": str(exc)}, ensure_ascii=False))
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
