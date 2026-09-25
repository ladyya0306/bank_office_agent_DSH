"""`db-entity-key`：可选主体识别码与证件资料。

缺统一社会信用代码、行内编号或身份证号只提示，不拦建档、填报和输出。
证件类型和号码按原文分别保存；不能把“证件号码”默认认作身份证号。
已提供的识别码仍查重，格式异常仅提示。只在 SCHEMA v2 上工作。
"""
from __future__ import annotations

import re
import sqlite3
from pathlib import Path

# 事件链与批次锚点的**唯一实现**在 store_v2（避免两处各写一套公式）。
# 这里保留同名再导出，是为了 selftest 与既有调用点不用改口径。
from .store_v2 import (  # noqa: F401  (re-exported)
    digest as _digest,
    open_db as _open_db,
    require_v2,
    resolve_batch,
    schema_version,
    write_event,
)

#: 识别键的种类。企业前两种，自然人第三种。
KEY_KINDS = ("uscc", "bank_no", "id_card")
KIND_LABEL = {"uscc": "统一社会信用代码", "bank_no": "行内编号",
              "id_card": "身份证号", "document": "其他证件号码"}

USCC_ALPHABET = "0123456789ABCDEFGHJKLMNPQRTUWXY"
USCC_WEIGHTS = (1, 3, 9, 27, 19, 26, 16, 17, 20, 29, 25, 13, 8, 24, 10, 30, 28)


class IdentityKeyError(RuntimeError):
    pass


def _now() -> str:
    """UTC 时刻 + 明确时区（`06 §2.2`，G-B8）——与 `store_v2` 同一口径。"""
    from .store_v2 import now_utc

    return now_utc()


# ---------------------------------------------------------------- format hints
def uscc_check_digit_ok(value: str) -> bool:
    """GB 32100 校验位。只在看起来像 USCC 时调用。"""
    v = value.strip().upper()
    if len(v) != 18 or any(c not in USCC_ALPHABET for c in v):
        return False
    total = sum(USCC_ALPHABET.index(c) * w for c, w in zip(v[:17], USCC_WEIGHTS))
    r = 31 - total % 31
    return USCC_ALPHABET[r % 31] == v[17]


def format_hint(kind: str, value: str) -> str | None:
    """返回一句"看起来不太像"的提示，或 None。**从不据此拒绝。**"""
    v = (value or "").strip()
    if not v:
        return "空值"
    if kind == "uscc":
        if len(v) == 18 and re.match(r"^[0-9A-Z]{18}$", v.upper()):
            return None if uscc_check_digit_ok(v) else "18 位，但**校验位算不过**（仍按你给的收下）"
        return "不是 18 位大写字母数字（仍按你给的收下）"
    if kind == "bank_no":
        return None if re.match(r"^\d{19}$", v) else "不是 19 位数字（本行账号是 19 位，仍按你给的收下）"
    if kind == "id_card":
        return None if re.match(r"^\d{17}[\dXx]$", v) else "不是 18 位身份证格式（仍按你给的收下）"
    return None


# ------------------------------------------------------------------ hash chain
# `write_event` / `resolve_batch` / `require_v2` 的实现在 `store_v2.py`（唯一一份公式）。
# 本模块只从那里 import，见文件头的再导出。


# ------------------------------------------------------------------- the gate
def _missing_kinds(row) -> list[str]:
    """仅供资料提示；已有其他类型证件号码的主体不算缺资料。"""
    if row["other_document"]:
        return []
    etype = row["entity_type"]
    if etype == "natural_person":
        return [] if (row["id_card"] or "").strip() else ["document"]
    have = [k for k in ("uscc", "bank_no") if (row[k] or "").strip()]
    return [] if have else ["uscc", "bank_no"]


def pending(con: sqlite3.Connection) -> list[dict]:
    """识别资料未提供的主体清单，仅供查看，不触发停工。"""
    con.row_factory = sqlite3.Row
    out = []
    for row in con.execute(
            "SELECT e.id, e.entity_type, e.uscc, e.bank_no, e.id_card, e.identity_incomplete,"
            " e.identity_pending_reason,"
            " EXISTS(SELECT 1 FROM entity_profile p WHERE p.entity_id=e.id"
            " AND p.field='证件号码' AND p.value IS NOT NULL AND p.value<>'') AS other_document,"
            " (SELECT name FROM entity_name n WHERE n.entity_id=e.id ORDER BY"
            "   CASE n.name_kind WHEN '现用名' THEN 0 ELSE 1 END, n.id LIMIT 1) AS name"
            " FROM entity e ORDER BY e.id"):
        miss = _missing_kinds(row)
        if not miss:
            continue
        out.append({"entity_id": row["id"], "entity_type": row["entity_type"],
                    "name": row["name"], "missing": miss,
                    "reason": row["identity_pending_reason"],
                    "flagged": bool(row["identity_incomplete"])})
    return out


def gate(con: sqlite3.Connection) -> dict:
    """识别资料完整性提示；缺失不阻止按工作区开展单户业务。"""
    require_v2(con)
    missing = pending(con)
    return {"ok": True, "blocked_count": 0, "blocked": [],
            "missing_count": len(missing), "missing": missing,
            "rule": "识别码、证件号码均为可选资料；缺失只提示，不阻止填报"}


def summarize(con: sqlite3.Connection) -> dict:
    require_v2(con)
    total = con.execute("SELECT COUNT(*) FROM entity").fetchone()[0]
    pend = pending(con)
    return {"entities": total, "pending": len(pend),
            "with_key": total - len(pend),
            "natural": con.execute("SELECT COUNT(*) FROM entity WHERE entity_type='natural_person'").fetchone()[0]}


# ------------------------------------------------------------------- mutations
def _find_by_key(con: sqlite3.Connection, kind: str, value: str) -> int | None:
    col = {"uscc": "uscc", "bank_no": "bank_no", "id_card": "id_card"}[kind]
    r = con.execute('SELECT id FROM entity WHERE "%s"=?' % col, (value.strip(),)).fetchone()
    return r[0] if r else None


def set_key(con: sqlite3.Connection, *, entity_id: int, kind: str, value: str,
            actor: str = "user", batch_no: str | None = None, algo: str = "sm3") -> dict:
    """给已有主体补一个识别键。**不接受空值**（空值等于没补）。"""
    require_v2(con)
    if kind not in KEY_KINDS:
        raise IdentityKeyError("未知的键类型 %r（只能是 %s）" % (kind, "/".join(KEY_KINDS)))
    value = (value or "").strip()
    if not value:
        raise IdentityKeyError("识别键不能为空——空值等于没补（26 §1.1.1）")
    row = con.execute("SELECT id, entity_type FROM entity WHERE id=?", (entity_id,)).fetchone()
    if row is None:
        raise IdentityKeyError("没有这个主体：id=%s" % entity_id)
    if row[1] == "natural_person" and kind != "id_card":
        raise IdentityKeyError("自然人主体的识别键是身份证号，不能用 %s" % KIND_LABEL[kind])
    if row[1] != "natural_person" and kind == "id_card":
        raise IdentityKeyError("企业主体的识别键是统一社会信用代码或行内编号，不能用身份证号")

    other = _find_by_key(con, kind, value)
    if other is not None and other != entity_id:
        raise IdentityKeyError(
            "这个%s已经属于主体 #%s——**不新建、也不覆盖**。"
            "要合并请走 db-merge 的勾选覆盖（26 §1.1.1）" % (KIND_LABEL[kind], other))

    con.execute('UPDATE entity SET "%s"=?, identity_incomplete=0, identity_pending_reason=NULL,'
                " updated_at=? WHERE id=?" % kind, (value, _now(), entity_id))
    bno = resolve_batch(con, entity_id, batch_no)
    hint = format_hint(kind, value)
    write_event(con, batch_no=bno, event_type="identity_key_set", actor_id=actor,
                entity_id=entity_id, target=KIND_LABEL[kind],
                payload={"entity_id": entity_id, "kind": kind, "value": value,
                         "source": "user", "hint": hint},
                value_text=value, algo=algo)
    con.commit()
    return {"entity_id": entity_id, "kind": kind, "label": KIND_LABEL[kind], "value": value,
            "batch_no": bno, "hint": hint,
            "still_pending": _missing_kinds(con.execute(
                "SELECT entity_type, uscc, bank_no, id_card,"
                " EXISTS(SELECT 1 FROM entity_profile p WHERE p.entity_id=entity.id"
                " AND p.field='证件号码' AND p.value IS NOT NULL AND p.value<>'')"
                " AS other_document FROM entity WHERE id=?",
                (entity_id,)).fetchone())}


def add_entity(con: sqlite3.Connection, *, name: str, kind: str | None = None, value: str | None = None,
               entity_type: str = "legal_person", actor: str = "user",
               batch_no: str | None = None, algo: str = "sm3") -> dict:
    """按名称新建主体；识别码可选，同时建立名称与案子。"""
    require_v2(con)
    name = (name or "").strip()
    if not name:
        raise IdentityKeyError("必须先有主体名")
    if kind is not None and kind not in KEY_KINDS:
        raise IdentityKeyError("未知的键类型 %r" % kind)
    value = (value or "").strip()
    if bool(kind) != bool(value):
        raise IdentityKeyError("给识别码时须同时提供类型和值；也可以都不填")
    dup = _find_by_key(con, kind, value) if kind else None
    if dup is not None:
        raise IdentityKeyError(
            "这个%s已存在（主体 #%s）——**不新建**，请改用 db-merge 的勾选覆盖" % (KIND_LABEL[kind], dup))
    if entity_type == "natural_person" and kind and kind != "id_card":
        raise IdentityKeyError("自然人不能用企业信用代码建档；可不填识别码，或用 --doc-type 保存其他证件")
    if entity_type != "natural_person" and kind == "id_card":
        raise IdentityKeyError("企业主体请用统一社会信用代码或行内编号作识别键")

    now = _now()
    cols = {"uscc": "uscc", "bank_no": "bank_no", "id_card": "id_card"}
    if kind:
        eid = con.execute(
            "INSERT INTO entity(entity_type,%s,identity_incomplete,created_at,updated_at)"
            " VALUES(?,?,0,?,?)" % cols[kind], (entity_type, value, now, now)).lastrowid
    else:
        eid = con.execute(
            "INSERT INTO entity(entity_type,identity_incomplete,created_at,updated_at)"
            " VALUES(?,1,?,?)", (entity_type, now, now)).lastrowid
    con.execute("INSERT INTO entity_name(entity_id,name,name_kind,created_at)"
                " VALUES(?,?,'现用名',?)", (eid, name, now))
    con.execute('INSERT INTO "case"(entity_id,created_at) VALUES(?,?)', (eid, now))
    bno = resolve_batch(con, eid, batch_no)
    hint = format_hint(kind, value) if kind else None
    write_event(con, batch_no=bno, event_type="identity_key_set", actor_id=actor,
                entity_id=eid, target=KIND_LABEL[kind] if kind else "主体建档",
                payload={"entity_id": eid, "kind": kind, "value": value, "source": "user",
                         "new_entity": True, "name": name, "hint": hint},
                value_text=value, algo=algo)
    con.commit()
    return {"entity_id": eid, "name": name, "kind": kind, "value": value,
            "batch_no": bno, "hint": hint}


def set_document(con: sqlite3.Connection, *, entity_id: int, doc_type: str,
                 doc_number: str | None = None, actor: str = "user",
                 algo: str = "sm3") -> dict:
    """可选证件资料按原文类型保存；不推断为身份证，也不强制给号码。"""
    require_v2(con)
    doc_type = (doc_type or "").strip()
    doc_number = (doc_number or "").strip()
    if not doc_type:
        raise IdentityKeyError("填写证件号码时请同时写明证件类型；证件资料也可以都不填")
    if con.execute("SELECT id FROM entity WHERE id=?", (entity_id,)).fetchone() is None:
        raise IdentityKeyError("没有这个主体：id=%s" % entity_id)
    now = _now()
    for field, value in (("证件类型", doc_type), ("证件号码", doc_number)):
        if not value:
            continue
        con.execute("INSERT INTO entity_profile(entity_id,field,value,created_at)"
                    " VALUES(?,?,?,?)", (entity_id, field, value, now))
    if doc_number:
        con.execute("UPDATE entity SET identity_incomplete=0,"
                    " identity_pending_reason=NULL, updated_at=? WHERE id=?",
                    (now, entity_id))
    bno = resolve_batch(con, entity_id, None)
    write_event(con, batch_no=bno, event_type="identity_key_set", actor_id=actor,
                entity_id=entity_id, target="可选证件资料",
                payload={"entity_id": entity_id, "doc_type": doc_type,
                         "doc_number": doc_number or None, "source": "user"}, algo=algo)
    con.commit()
    return {"entity_id": entity_id, "doc_type": doc_type,
            "doc_number": doc_number or None, "batch_no": bno}


def note_unavailable(con: sqlite3.Connection, *, entity_id: int, reason: str,
                     actor: str = "user", algo: str = "sm3") -> dict:
    """可选地记录暂时没有识别码的原因；不会阻断工作。"""
    require_v2(con)
    row = con.execute("SELECT id FROM entity WHERE id=?", (entity_id,)).fetchone()
    if row is None:
        raise IdentityKeyError("没有这个主体：id=%s" % entity_id)
    reason = (reason or "").strip()
    if not reason:
        raise IdentityKeyError("必须写明原因（这是给以后的人看的）")
    con.execute("UPDATE entity SET identity_incomplete=1, identity_pending_reason=? WHERE id=?",
                (reason, entity_id))
    bno = resolve_batch(con, entity_id, None)
    write_event(con, batch_no=bno, event_type="identity_key_requested", actor_id=actor,
                entity_id=entity_id, target="识别键",
                payload={"entity_id": entity_id, "reason": reason, "outcome": "unavailable",
                         "effect": "informational —— 不阻断工作"},
                algo=algo)
    con.commit()
    return {"entity_id": entity_id, "recorded": reason, "still_blocked": False}


def note_requested(con: sqlite3.Connection, *, actor: str = "user", algo: str = "sm3") -> dict:
    """仅在用户主动索取识别资料时记一条请求；日常流程不自动调用。"""
    require_v2(con)
    pend = pending(con)
    if not pend:
        return {"requested": 0, "items": []}
    by_batch: dict[str, list] = {}
    for p in pend:
        by_batch.setdefault(resolve_batch(con, p["entity_id"], None), []).append(p)
    for bno, items in by_batch.items():
        write_event(con, batch_no=bno, event_type="identity_key_requested", actor_id=actor,
                    target="识别键",
                    payload={"count": len(items),
                             "entities": [{"entity_id": i["entity_id"], "name": i["name"],
                                           "missing": i["missing"]} for i in items]},
                    algo=algo)
    con.commit()
    return {"requested": len(pend), "items": pend}


# --------------------------------------------------------------- human report
def render_report(kind: str, info: dict, res: dict | None = None) -> str:
    L: list[str] = []
    A = L.append
    if kind in ("list", "gate"):
        A("# 可选识别资料检查")
        A("")
        A("识别码和证件号码不是开工条件；缺失只作为资料提示。")
        A("")
        A("| 项 | 数 |")
        A("| --- | --- |")
        A("| 主体总数 | %d |" % info["entities"])
        A("| **已有识别键** | **%d** |" % info["with_key"])
        A("| 未提供常见识别码（不阻断） | %d |" % info["pending"])
        A("")
        if info["pending"]:
            A("## 未提供常见识别码的主体")
            A("")
            A("| # | 主体 | 类型 | 缺什么 | 备注 |")
            A("| --- | --- | --- | --- | --- |")
            for i in info["missing"]:
                A("| %s | %s | %s | %s | %s |" % (
                    i["entity_id"], i["name"] or "（无名）",
                    "自然人" if i["entity_type"] == "natural_person" else "企业",
                    "、".join(KIND_LABEL[k] for k in i["missing"]),
                    i["reason"] or "—"))
            A("")
            A("## 结论：可以继续工作")
            A("")
            A("如有资料，可使用 `office.py db-entity-key` 补录；没有也不影响填报。")
        else:
            A("## 结论：可以继续工作")
    else:
        A("# 识别键补录结果")
        A("")
        A("| 项 | 值 |")
        A("| --- | --- |")
        for k, v in (res or {}).items():
            A("| %s | %s |" % (k, v))
        A("")
        if (res or {}).get("hint"):
            A("> ⚠️ 格式提示：%s" % res["hint"])
            A("> 仅作格式提示，不阻断工作。")
    return "\n".join(L) + "\n"


# ------------------------------------------------------------------ CLI glue
def open_v2(db_path: str | Path) -> sqlite3.Connection:
    p = Path(db_path)
    if not p.exists():
        raise IdentityKeyError("数据库不存在：%s" % p)
    con = sqlite3.connect(str(p))
    con.row_factory = sqlite3.Row
    try:
        require_v2(con)
    except Exception as exc:  # noqa: BLE001 - 统一成 IdentityKeyError，命令层好读
        con.close()
        raise IdentityKeyError(str(exc)) from exc
    return con


def cmd_db_entity_key(args):  # pragma: no cover - exercised by selftest
    from . import workroot as WR
    from .common import OfficeKitError, Result, write_text, unique_path

    try:
        WR.open_for_command(args)
    except WR.WorkrootError as exc:
        raise OfficeKitError(str(exc)) from exc
    con = open_v2(args._db)
    out = getattr(args, "out", None)
    res = Result("db-entity-key")
    try:
        eid = getattr(args, "entity_id", None)

        if getattr(args, "add", False):
            if getattr(args, "doc_number", None) and not getattr(args, "doc_type", None):
                raise IdentityKeyError("填写证件号码时请同时写明证件类型")
            kind, value = _pick_key(args, optional=True)
            info = add_entity(con, name=args.name, kind=kind, value=value,
                              entity_type=getattr(args, "entity_type", "legal_person"),
                              actor=getattr(args, "by", "user") or "user")
            if getattr(args, "doc_type", None) or getattr(args, "doc_number", None):
                info["document"] = set_document(
                    con, entity_id=info["entity_id"],
                    doc_type=getattr(args, "doc_type", None),
                    doc_number=getattr(args, "doc_number", None),
                    actor=getattr(args, "by", "user") or "user")
            text = render_report("set", info)
            res.data.update({"action": "add", **info})
            if info.get("hint"):
                res.warn("识别键格式提示：%s（只是提示，不拒绝）" % info["hint"])

        elif getattr(args, "unavailable", False):
            info = note_unavailable(con, entity_id=eid, reason=args.reason,
                                    actor=getattr(args, "by", "user") or "user")
            text = render_report("set", {"动作": "记下'拿不到'", "主体": eid,
                                         "原因": args.reason,
                                         "结果": "已记录；不影响继续工作"})
            res.data.update({"action": "unavailable", **info})

        elif eid is not None:
            kind, value = _pick_key(args, optional=True)
            if kind:
                info = set_key(con, entity_id=eid, kind=kind, value=value,
                               actor=getattr(args, "by", "user") or "user",
                               batch_no=getattr(args, "batch", None))
            elif getattr(args, "doc_type", None) or getattr(args, "doc_number", None):
                info = set_document(con, entity_id=eid,
                                    doc_type=getattr(args, "doc_type", None),
                                    doc_number=getattr(args, "doc_number", None),
                                    actor=getattr(args, "by", "user") or "user")
            else:
                raise IdentityKeyError("请填写一个识别码，或填写 --doc-type 及可选的 --doc-number")
            text = render_report("set", info)
            res.data.update({"action": "set", **info})
            if info.get("hint"):
                res.warn("识别键格式提示：%s（只是提示，不拒绝）" % info["hint"])

        else:
            # 只展示可选资料的缺失情况，不自动要求用户补齐。
            g = gate(con)
            g.update(summarize(con))
            g["requested"] = 0
            text = render_report("gate", g)
            res.data.update({"action": "gate", **g})
            if g["missing_count"]:
                res.warn("%d 个主体未提供识别码或身份证号；这些资料可选。"
                         % g["missing_count"])

        if out or getattr(args, "_workroot", None):
            bno = None
            try:
                row = con.execute("SELECT batch_no FROM source WHERE batch_no IS NOT NULL"
                                  " ORDER BY id DESC LIMIT 1").fetchone()
                bno = row[0] if row else None
            except Exception:  # noqa: BLE001
                bno = None
            d = Path(out) if out else WR.report_dir_for(args, "db-entity-key", bno)
            p = write_text(d / "identifier_report.md", text)
            res.add_artifact(p, "识别键检查/补录报告")
        res.data["report"] = text
        return res
    finally:
        con.close()


def _pick_key(args, *, optional: bool = False):
    given = [(k, getattr(args, k, None)) for k in KEY_KINDS if getattr(args, k, None)]
    if not given:
        if optional:
            return None, None
        raise IdentityKeyError("要给一个识别键：--uscc / --bank-no / --id-card 三选一")
    if len(given) > 1:
        raise IdentityKeyError("一次只补一个键（同时给了 %s）" % "、".join(k for k, _ in given))
    return given[0]
