"""`db-merge`——同键勾选覆盖（**T-19**）。

规则（唯一权威：`27 §5.3`、`26 §3.1`）

* 库里已经有同一个键、值却不一样 → **出可勾选表**，**你不勾就不动**。
* 勾了 → **写一条新 `fact` 行**，旧行标 `superseded_by`（历史不丢，`17 §3.8`），
  并写 `fact_superseded` 事件。
* **已交付的产物不受影响**：文件是死的，覆盖数据库**不会**回头改已经出去的 Word。
  所以本模块对这类字段**默认拒绝勾选**，要覆盖必须显式加 `--include-delivered`，
  并在事件里记下"这份产物与库已经不一致"。

源侧四选与勾选动作的对应（`27 §5.1` 甲表，**逐条对上**）：

| 源侧选择 | 本命令的勾选方式 | 写什么 |
| --- | --- | --- |
| 接受入库 | `--select N` | 覆盖为新值，`source_kind` 不变 |
| 改后入库 | `--new N=你输入的值` | 覆盖，`source_kind='user'`、`provenance='你确认的'` |
| **丢弃** | `--discard N` | **不写库**，但 🔴 写 `fact_discarded` 事件 |
| 留空待查 | `--defer N` | 不动，冲突保持未裁定、进 `review_queue` |
| （另外）我看过了、保留原值 | `--keep N` | 不动，裁定为 `took_existing` |

⚠️ **`--select-all` 不覆盖两类行**（这是对"高风险五类无论置信度多高都必须过使用者复核"
[`26 §2.5`] 的落实，也是本工具的一处**实现口径**，写在报告里让评审看得见）：
**高风险五类字段**（收款账号 / 金额 / 利率 / 日期 / 证件号码）与**已经交付过的字段**，
都必须**逐个**写编号勾。
"""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

from . import doc_fill
from .common import OfficeKitError, Result, out_dir, unique_path, write_text
from . import workroot as WR
from .store_v2 import StoreV2, resolve_batch

#: 与 `migrate_v2` 同源，避免两处各写一套"高风险"判断
from .migrate_v2 import HIGH_RISK_RULES  # noqa: E402


class MergeError(RuntimeError):
    pass


def high_risk_class(field: str) -> str | None:
    """属于高风险五类的哪一类；不属于返回 None（`26 §2.6`）。"""
    for label, words in HIGH_RISK_RULES:
        if any(w in field for w in words):
            return label
    return None


# --------------------------------------------------------------------------
# plan
# --------------------------------------------------------------------------
def _delivered_index(store: StoreV2) -> dict[str, list[dict]]:
    idx = store.delivered_fields()
    out: dict[str, list[dict]] = {}
    for field, rows in idx.items():
        keep = [r for r in rows if r.get("artifact_path")]
        if keep:
            out[field] = keep
    return out


def _entity_hint(profile: dict[str, dict[str, Any]]) -> str | None:
    """从 incoming 里挑一个可以作为归属锚点的主体名。

    只认**主体名字段**（`schema_v2.SUBJECT_NAME_KEYS`）——不猜、不从别的键里凑。
    """
    for key in ("借款人名称", "保证人名称", "借款人法定代表人", "保证人法定代表人",
                "个人保证人姓名"):
        meta = profile.get(key) or {}
        if meta.get("value"):
            return str(meta["value"])
    return None


def _scope_entity(store: StoreV2, key: str, default: int | None) -> tuple[int | None, bool]:
    """这一行该归属哪个主体。

    * 给了主体（`--entity-id` 或从主体名字段推出来）→ 就用它；
    * 没给 → 看库里这个键**当前**属于谁：只有一个主体 → 就是它；
      **多个主体都有这个键 → 不猜**，交给使用者用 `--entity-id` 指明（返回 ambiguous=True）；
    * 库里根本没有 → 未归属（None）。
    """
    if default is not None:
        return default, False
    rows = [r[0] for r in store.conn.execute(
        "SELECT DISTINCT entity_id FROM fact WHERE key=? AND superseded_by IS NULL"
        " AND entity_id IS NOT NULL", (key,))]
    if len(rows) == 1:
        return int(rows[0]), False
    return None, len(rows) > 1


def build_plan(store: StoreV2, *, incoming: dict[str, dict[str, Any]] | None = None,
               entity_id: int | None = None, batch_no: str | None = None,
               include_new: bool = False, stage: str = "plan") -> dict:
    """把"待裁定的同键差异"整理成一份可勾选清单。

    两个来源，产出同一种行：

    * `incoming` 给了 → 与库里现值逐个比对（**新键只报告、不入库**，入库请用 `db-ingest`）
    * `incoming` 没给 → 直接用库里未裁定的 `fact_conflict` 行

    ⚠️ **新材料没给值的键不进勾选表**：那不是"值不一样"，是"这次没带"。
    拿空值去覆盖库里已有的值 = 把已知信息抹掉，所以只在报告里单列一段（`incoming_blank`）。
    """
    delivered = _delivered_index(store)
    rows: list[dict] = []
    blank: list[dict] = []

    if incoming is None:
        # 来源一：库里已经记下的冲突
        for c in store.conflicts(open_only=True):
            key = c["key"]
            eid = c["entity_id"]
            if entity_id is not None and c["entity_id"] != entity_id:
                continue  # 只裁这个主体的
            rows.append({
                "key": key, "entity_id": eid, "entity_name": store.entity_label(eid),
                "existing": c["existing"], "incoming": c["incoming"],
                "existing_source_id": c["existing_src"], "incoming_source_id": c["incoming_src"],
                "conflict_id": c["id"], "batch_no": c["batch_no"],
                "norm_rule": c["norm_rule"],
                "source_kind": "source", "provenance": None,
                "origin": "库里待裁定的冲突",
                "kind": "conflict",
            })
    else:
        hint = _entity_hint(incoming)
        base = entity_id
        if base is None and hint:
            base = store.find_entity_by_name(hint)
        for key, meta in incoming.items():
            text = None if meta.get("missing") else (
                None if meta.get("value") is None else str(meta["value"]))
            eid, ambiguous = _scope_entity(store, key, base)
            cur = store.current_row(key, eid)
            prov = meta.get("origin") or meta.get("source")
            if text is None or not str(text).strip():
                if cur is not None:
                    blank.append({"key": key, "entity_id": eid,
                                  "entity_name": store.entity_label(eid),
                                  "existing": cur["value"], "origin": prov})
                continue
            if cur is None:
                if not include_new:
                    continue
                rows.append({
                    "key": key, "entity_id": eid, "entity_name": store.entity_label(eid),
                    "existing": None, "incoming": text,
                    "existing_source_id": None, "incoming_source_id": None,
                    "conflict_id": None, "batch_no": batch_no, "norm_rule": None,
                    "source_kind": "source", "provenance": prov,
                    "origin": prov, "kind": "new", "scope_warning": None,
                })
                continue
            if (cur["value"] or "") == (text or ""):
                continue  # 没有差异，不该出现在勾选表里
            was_blank = not (cur["value"] or "").strip()
            rows.append({
                "key": key, "entity_id": eid, "entity_name": store.entity_label(eid),
                "existing": cur["value"], "incoming": text,
                "existing_source_id": cur["source_id"], "incoming_source_id": None,
                "conflict_id": None, "batch_no": batch_no, "norm_rule": None,
                "source_kind": "source", "provenance": prov,
                "origin": prov or "（未注明出处）",
                # 库里原来是空的 → 这是"把空白补上"，**不是覆盖**，丢不了东西
                "kind": "fill_blank" if was_blank else "conflict",
                "scope_warning": ("库里多个主体都有这个键，且这次没指明主体——"
                                  "请加 --entity-id 再出一次清单" if ambiguous else None),
            })

    for i, row in enumerate(rows, start=1):
        row["n"] = i
        row.setdefault("scope_warning", None)
        row.setdefault("kind", "conflict")
        row["high_risk"] = high_risk_class(row["key"])
        row["delivered"] = [
            {"artifact_path": d["artifact_path"], "sha256": d.get("artifact_sha256"),
             "signoff": d.get("signoff_decision"), "at": d.get("created_at")}
            for d in delivered.get(row["key"], [])
        ]
        # "补空白"丢不了东西，所以不需要逐个点头；高风险与已交付字段仍然要
        row["must_tick_individually"] = bool(
            row["high_risk"] or row["delivered"] or row["scope_warning"]) \
            and row["kind"] != "fill_blank"

    return {
        "stage": stage,
        "batch_no": batch_no,
        "entity_id": entity_id,
        "rows": rows,
        "incoming_blank": blank,
        "counts": {
            "rows": len(rows),
            "conflict": sum(1 for r in rows if r["kind"] == "conflict"),
            "fill_blank": sum(1 for r in rows if r["kind"] == "fill_blank"),
            "new": sum(1 for r in rows if r["kind"] == "new"),
            "high_risk": sum(1 for r in rows if r["high_risk"]),
            "delivered": sum(1 for r in rows if r["delivered"]),
            "incoming_blank": len(blank),
            "ambiguous_scope": sum(1 for r in rows if r["scope_warning"]),
            "entity_scoped": entity_id is not None,
        },
    }


# --------------------------------------------------------------------------
# selection
# --------------------------------------------------------------------------
def parse_selection(select: str | None, select_all: bool = False,
                    new: list[str] | None = None, discard: list[str] | None = None,
                    defer: list[str] | None = None, keep: list[str] | None = None,
                    include_delivered: bool = False) -> dict:
    """把命令行上的勾选解析成 `{编号: 动作}`。**动作之间不许重复。**"""
    actions: dict[int, dict] = {}

    def take(num: str, action: str, value: str | None = None) -> None:
        try:
            n = int(str(num).split("=")[0].strip())
        except ValueError as exc:
            raise MergeError("勾选编号必须是数字：%r" % num) from exc
        if n in actions:
            raise MergeError("编号 %d 被勾了两次（%s 与 %s）——一次只能给一个动作"
                             % (n, actions[n]["action"], action))
        actions[n] = {"action": action, "value": value}

    for token in (select or "").replace("，", ",").split(","):
        if token.strip():
            take(token, "accept")
    for token in new or []:
        if "=" not in str(token):
            raise MergeError("--new 要写成 编号=值，例如 --new 3=广东众森实业发展有限公司")
        num, value = str(token).split("=", 1)
        take(num, "new", value)
    for token in discard or []:
        take(token, "discard")
    for token in defer or []:
        take(token, "defer")
    for token in keep or []:
        take(token, "keep")

    chosen = set(actions)
    return {"actions": actions, "select_all": bool(select_all), "chosen": sorted(chosen),
            "include_delivered": bool(include_delivered)}


def _guard(plan: dict, selection: dict) -> tuple[list[dict], list[dict]]:
    """把勾选与计划对上；不合法的一律**拒绝并说明**，绝不悄悄跳过。"""
    rows = {r["n"]: r for r in plan["rows"]}
    accepted: list[dict] = []
    problems: list[dict] = []
    unknown = [n for n in selection["chosen"] if n not in rows]
    for n in unknown:
        problems.append({"n": n, "why": "计划里没有这个编号"})

    todo: dict[int, dict] = {}
    if selection["select_all"]:
        for r in plan["rows"]:
            if r["must_tick_individually"]:
                why = []
                if r["high_risk"]:
                    why.append("高风险五类（%s）" % r["high_risk"])
                if r["delivered"]:
                    why.append("已交付过 %d 份产物" % len(r["delivered"]))
                problems.append({"n": r["n"], "why": "全选不含它：" + "、".join(why) + "，请单独写编号"})
                continue
            if r["kind"] == "new":
                continue  # 新键没有"覆盖"，全选不管它
            todo[r["n"]] = {"action": "accept", "value": None}
    for n, act in selection["actions"].items():
        if n not in rows:
            continue
        todo[n] = act

    for n in sorted(todo):
        r, act = rows[n], todo[n]
        if r["scope_warning"]:
            problems.append({"n": n, "why": r["scope_warning"]})
            continue
        # "已交付"只挡**会改值**的动作（accept / new）。丢弃、留空待查、保留原值
        # 都不动库里当前那个值，也就动不到已经出去的产物，没必要挡。
        if r["delivered"] and act["action"] in ("accept", "new") \
                and not selection["include_delivered"]:
            problems.append({"n": n, "why":
                             "这个字段已经填进过产物（%s）。覆盖数据库**不会**改已经出去的文件，"
                             "真要覆盖请加 --include-delivered" % r["delivered"][0]["artifact_path"]})
            continue
        accepted.append({**r, **act})
    return accepted, problems


# --------------------------------------------------------------------------
# apply
# --------------------------------------------------------------------------
def apply_plan(store: StoreV2, plan: dict, selection: dict, *, actor: str = "user",
               batch_no: str | None = None) -> dict:
    """按勾选落库。**每条都写事件，四种动作各有各的记法。**"""
    accepted, problems = _guard(plan, selection)
    if problems and not accepted:
        return {"applied": 0, "problems": problems, "results": []}

    results: list[dict] = []
    for row in accepted:
        bno = batch_no or row.get("batch_no") or plan.get("batch_no") or \
            resolve_batch(store.conn, row["entity_id"], None)
        key, eid, act = row["key"], row["entity_id"], row["action"]
        # 计划过期检查：库里现值必须还是计划里那个"现值"
        cur = store.current_row(key, eid)
        cur_value = None if cur is None else cur["value"]
        if (cur_value or "") != (row["existing"] or ""):
            problems.append({"n": row["n"], "why":
                             "计划已过期：计划里库中现值是 %r，现在变成了 %r。请重新出清单"
                             % (row["existing"], cur_value)})
            continue

        if act == "discard":
            store.discard_fact(bno, key, row["incoming"], entity_id=eid,
                               reason="源侧复核：丢弃（db-merge）")
            if row.get("conflict_id"):
                store.resolve_conflict(row["conflict_id"], "left_blank", decided_by=actor,
                                       batch_no=bno)
            results.append({"n": row["n"], "key": key, "action": "discarded",
                            "old": row["existing"], "new": None})
            continue

        if act == "defer":
            store.add_review(bno, "field", "merge_deferred", entity_id=eid, field=key,
                             label=key, candidates={"existing": row["existing"],
                                                    "incoming": row["incoming"]})
            results.append({"n": row["n"], "key": key, "action": "deferred",
                            "old": row["existing"], "new": row["incoming"]})
            continue

        if act == "keep":
            if row.get("conflict_id"):
                store.resolve_conflict(row["conflict_id"], "took_existing", decided_by=actor,
                                       batch_no=bno)
            store.event("fact_conflict_detected", batch_no=bno, entity_id=eid, target=key,
                        payload={"key": key, "outcome": "took_existing",
                                 "kept": row["existing"], "rejected": row["incoming"],
                                 "decided_by": actor,
                                 "why": "使用者看过清单后选择保留库里原值"},
                        value_text=row["existing"])
            results.append({"n": row["n"], "key": key, "action": "kept",
                            "old": row["existing"], "new": row["existing"]})
            continue

        # accept / new → 覆盖或新增
        if act == "new":
            value = row["value"]
            source_kind, provenance = "user", "你确认的"
        else:
            value = row["incoming"]
            source_kind, provenance = row["source_kind"], row["provenance"]
        if row["kind"] == "new":
            out = store.put_fact(bno, key, value, entity_id=eid,
                                 source_kind=source_kind, provenance=provenance,
                                 status="ok", on_conflict="ignore")
            action = "added"
        else:
            out = store.put_fact(bno, key, value, entity_id=eid,
                                 source_kind=source_kind, provenance=provenance,
                                 status="ok", on_conflict="supersede")
            action = "superseded"
            if row.get("conflict_id"):
                store.resolve_conflict(row["conflict_id"], "took_incoming", decided_by=actor,
                                       batch_no=bno)
        if row["delivered"]:
            store.event("fact_superseded", batch_no=bno, entity_id=eid, target=key,
                        payload={"key": key, "entity_id": eid,
                                 "effect": "🔴 已交付产物与库不再一致",
                                 "artifacts": [d["artifact_path"] for d in row["delivered"]],
                                 "old_value": row["existing"], "new_value": value},
                        value_text=value)
        results.append({"n": row["n"], "key": key, "action": action,
                        "old": row["existing"], "new": value,
                        "fact_id": out.get("fact_id"), "decided_by": actor})

    return {"applied": len(results), "problems": problems, "results": results}


# --------------------------------------------------------------------------
# 报告
# --------------------------------------------------------------------------
def render_plan(plan: dict) -> str:
    L: list[str] = []
    A = L.append
    A("# 同键勾选覆盖清单 / db-merge")
    A("")
    A("库里已经有同一个键、但值不一样。**按编号勾，不勾就不动。**")
    A("")
    A("| 项 | 数 |")
    A("| --- | --- |")
    A("| 待裁定行 | **%d** |" % plan["counts"]["rows"])
    A("| 其中：冲突（库里已有值） | %d |" % plan["counts"]["conflict"])
    A("| 其中：库里原来是空的（**补空白，不是覆盖**） | %d |" % plan["counts"]["fill_blank"])
    A("| 其中：库里还没有的键 | %d |" % plan["counts"]["new"])
    A("| 🔴 高风险五类字段（必须逐个勾） | %d |" % plan["counts"]["high_risk"])
    A("| ⚠️ 已交付过产物的字段（必须逐个勾） | %d |" % plan["counts"]["delivered"])
    A("| 本次材料没给值的键（**不进勾选表**） | %d |" % plan["counts"]["incoming_blank"])
    A("")
    if plan["counts"]["ambiguous_scope"]:
        A("🔴 **有 %d 行库里多个主体都有这个键，而这次没指明主体**——"
          "请加 `--entity-id <编号>` 再出一次清单，本工具**不替你猜**。"
          % plan["counts"]["ambiguous_scope"])
        A("")
    if not plan["rows"]:
        A("## 没有需要裁定的同键差异 ✅")
        A("")
        if plan.get("incoming_blank"):
            _blank_section(A, plan)
        return "\n".join(L) + "\n"

    A("## 勾选表")
    A("")
    A("| 编号 | 字段 | 主体 | 库里现值 | 新来的值 | 出处 | 标记 |")
    A("| --- | --- | --- | --- | --- | --- | --- |")
    for r in plan["rows"]:
        marks = []
        if r["high_risk"]:
            marks.append("🔴 高风险·%s" % r["high_risk"])
        if r["delivered"]:
            marks.append("⚠️ 已交付 %d 份" % len(r["delivered"]))
        if r["kind"] == "new":
            marks.append("➕ 库里没有")
        if r["kind"] == "fill_blank":
            marks.append("⬜ 库里原来是空的（填上不丢东西）")
        A("| %d | `%s` | %s | %s | %s | %s | %s |" % (
            r["n"], r["key"], r["entity_name"],
            _cell(r["existing"]), _cell(r["incoming"]),
            r["origin"] or "—",
            "、".join(marks + (["🔴 主体不明"] if r["scope_warning"] else [])) or "—"))
    A("")
    if plan.get("incoming_blank"):
        _blank_section(A, plan)
    A("## 怎么勾（**源侧四选，逐条对上**）")
    A("")
    A("```")
    A("--select 1,3            # 接受新值，覆盖入库")
    A("--new 2=你输入的值       # 改后入库（记 source_kind='user'、provenance='你确认的'）")
    A("--discard 5             # 丢弃：不写库，但 🔴 写 fact_discarded 事件")
    A("--defer 4               # 留空待查：不动，进 review_queue")
    A("--keep 6                # 看过了，保留库里原值（裁定 took_existing）")
    A("--select-all            # 全选（**不含**高风险五类与已交付字段——那两类要逐个勾）")
    A("--include-delivered     # 允许勾已交付的字段（默认拒绝）")
    A("```")
    A("")
    A("第二步把清单原样传回来，编号与库里现值会**再核对一遍**；对不上就拒绝执行（计划过期）。")
    A("")
    A("> ⚠️ **已交付的产物不受影响**：文件是死的，覆盖数据库**不会**回头改已经出去的 Word。")
    A("> 所以勾了已交付字段，事件里会明确记下「🔴 已交付产物与库不再一致」。")
    if plan["counts"]["high_risk"] or plan["counts"]["delivered"]:
        A("")
        A("> 🔴 **本工具的实现口径**：`--select-all` 不覆盖**高风险五类字段**"
          "（收款账号 / 金额 / 利率 / 日期 / 证件号码，[26 §2.5]）与**已交付字段**，"
          "这两类必须逐个写编号勾。这是对「高风险字段无论置信度多高都必须过使用者复核」的落实。")
    return "\n".join(L) + "\n"


def _blank_section(A, plan: dict) -> None:
    """本次材料没给值的键——**单列一段，不进勾选表**。"""
    A("## 本次材料没给这些键的值（库里原值**原样保留**）")
    A("")
    A("拿空值去覆盖已知值，等于把已经掌握的信息抹掉，所以这些行**不进勾选表**。")
    A("")
    A("| 字段 | 主体 | 库里现值 | 本次出处 |")
    A("| --- | --- | --- | --- |")
    for b in plan["incoming_blank"]:
        A("| `%s` | %s | %s | %s |" % (b["key"], b["entity_name"],
                                       _cell(b["existing"]), b["origin"] or "—"))
    A("")


def _cell(value: Any) -> str:
    if value is None or str(value) == "":
        return "（空）"
    text = str(value).replace("|", "\\|")
    return text if len(text) <= 40 else text[:37] + "…"


def render_applied(info: dict, plan: dict) -> str:
    L = ["# 同键勾选覆盖 · 执行结果", ""]
    rows = {r["n"]: r for r in plan.get("rows", [])}
    A = L.append
    A("| 项 | 数 |")
    A("| --- | --- |")
    A("| 本次落库 | **%d** |" % info["applied"])
    A("| 被拒绝（含计划过期） | %d |" % len(info["problems"]))
    A("")
    if info["results"]:
        A("## 落库明细")
        A("")
        A("| 编号 | 字段 | 动作 | 原值 | 现值 |")
        A("| --- | --- | --- | --- | --- |")
        label = {"superseded": "覆盖（旧行标 superseded_by）", "added": "新增",
                 "discarded": "丢弃（只留痕）", "deferred": "留空待查", "kept": "保留原值"}
        for r in info["results"]:
            A("| %d | `%s` | %s | %s | %s |" % (
                r["n"], r["key"], label.get(r["action"], r["action"]),
                _cell(r["old"]), _cell(r["new"])))
        A("")
    if info["problems"]:
        A("## 🔴 被拒绝的行（**没有执行，也没有悄悄跳过**）")
        A("")
        A("| 编号 | 原因 |")
        A("| --- | --- |")
        for p in info["problems"]:
            A("| %s | %s |" % (p["n"], p["why"]))
        A("")
    return "\n".join(L) + "\n"


# --------------------------------------------------------------------------
# CLI
# --------------------------------------------------------------------------
def cmd_db_merge(args) -> Result:
    res = Result("db-merge")
    try:
        wr = WR.open_for_command(args)
    except WR.WorkrootError as exc:
        raise OfficeKitError(str(exc)) from exc
    try:
        store = StoreV2(args._db, actor=getattr(args, "by", None) or "user")
    except Exception as exc:  # noqa: BLE001
        raise OfficeKitError(str(exc)) from exc
    out = WR.report_dir_for(args, "db-merge",
                            batch_no=getattr(args, "batch", None)
                            or store.current_batch())
    for _w in WR.outside_warnings(args, out):
        res.warn(_w)
    try:
        incoming = None
        inc_path = getattr(args, "incoming", None)
        if inc_path:
            incoming = doc_fill.load_profile(inc_path)

        entity_id = getattr(args, "entity_id", None)
        if entity_id is None:
            from .entity_key import KIND_LABEL

            for kind in ("uscc", "bank_no", "id_card"):
                value = getattr(args, kind, None)
                if not value:
                    continue
                entity_id = store.find_entity_by_key(kind, value)
                if entity_id is None:
                    raise OfficeKitError(
                        "库里没有用这个%s登记过的主体——**不新建**，"
                        "请先用 `db-entity-key --add --name <名> --%s <值>` 建档"
                        % (KIND_LABEL[kind], kind.replace("_", "-")))
                break
        if entity_id is None and getattr(args, "name", None):
            entity_id = store.find_entity_by_name(args.name)
            if entity_id is None:
                raise OfficeKitError("库里没有叫「%s」的主体——**不新建**，"
                                     "请先用 db-entity-key --add 建档" % args.name)
        if entity_id is None and incoming is not None:
            hint = _entity_hint(incoming)
            if hint:
                entity_id = store.find_entity_by_name(hint)
        if entity_id is not None:
            row = store.conn.execute("SELECT id FROM entity WHERE id=?", (entity_id,)).fetchone()
            if row is None:
                raise OfficeKitError("没有这个主体：id=%s" % entity_id)

        plan_path = getattr(args, "plan", None)
        batch_no = getattr(args, "batch", None)

        # ---- 第二步：从清单执行 ------------------------------------------
        if plan_path:
            p = Path(plan_path)
            if not p.exists():
                raise OfficeKitError("清单文件不存在：%s" % p)
            plan = json.loads(p.read_text(encoding="utf-8"))
            if not any([getattr(args, "select", None), getattr(args, "select_all", False),
                        getattr(args, "new", None), getattr(args, "discard", None),
                        getattr(args, "defer", None), getattr(args, "keep", None)]):
                raise OfficeKitError(
                    "要从清单执行，必须给出勾选（如 --select 1,3 / --select-all / --new 2=值）")
        else:
            # ---- 第一步：出清单 -------------------------------------------
            if incoming is None and entity_id is None and not store.conflicts(open_only=True):
                res.data.update({"stage": "plan", "rows": [], "counts": {"rows": 0}})
                res.warn("库里没有待裁定的同键差异，也没有给 --incoming——没什么可勾的")
                return res
            plan = build_plan(store, incoming=incoming, entity_id=entity_id,
                              batch_no=batch_no, include_new=getattr(args, "add_new", False))

        selecting = any([getattr(args, "select", None), getattr(args, "select_all", False),
                         getattr(args, "new", None), getattr(args, "discard", None),
                         getattr(args, "defer", None), getattr(args, "keep", None)])

        if not selecting:
            md = render_plan(plan)
            p_md = write_text(unique_path(out / "merge_plan.md"), md)
            p_json = write_text(unique_path(out / "merge_plan.json"),
                                json.dumps(plan, ensure_ascii=False, indent=2))
            res.add_artifact(p_md, "可勾选的同键覆盖清单（给使用者看）")
            res.add_artifact(p_json, "同一份清单（第二步用 --plan 传回来）")
            res.data.update({"stage": "plan", **plan, "plan_file": str(p_json)})
            if plan["counts"]["rows"]:
                res.warn(
                    "有 %d 个键库里已有不同的值，**没有覆盖**。请把 merge_plan.md 给使用者，"
                    "按编号勾选后用 `--plan %s --select 1,3` 执行（不勾就不动）"
                    % (plan["counts"]["rows"], p_json.name))
            else:
                res.warn("没有同键差异，库里的值就是最新的")
            return res

        # ---- 有勾选 → 执行 ------------------------------------------------
        selection = parse_selection(
            getattr(args, "select", None), getattr(args, "select_all", False),
            getattr(args, "new", None), getattr(args, "discard", None),
            getattr(args, "defer", None), getattr(args, "keep", None),
            getattr(args, "include_delivered", False))
        if not selection["chosen"] and not selection["select_all"]:
            raise OfficeKitError("给了勾选参数却一个编号都没有")
        info = apply_plan(store, plan, selection, actor=getattr(args, "by", None) or "user",
                          batch_no=batch_no)
        md = render_applied(info, plan)
        p_md = write_text(out / "merge_applied.md", md)
        res.add_artifact(p_md, "勾选覆盖的执行结果")
        res.data.update({"stage": "apply", "applied": info["applied"],
                         "problems": info["problems"], "results": info["results"],
                         "summary": store.summary(),
                         "plan_counts": plan["counts"]})
        if info["problems"]:
            res.warn("🔴 %d 行被拒绝（见报告 %s）" % (len(info["problems"]), p_md.name))
        if info["applied"]:
            res.warn("已按勾选覆盖 %d 个键；旧行仍留在库里并标了 superseded_by（历史不丢）"
                     % info["applied"])
        return res
    finally:
        store.close()
