"""SCHEMA v2 的运行时访问层——管线**唯一**使用的数据库入口。

为什么要有这一层
----------------
v1 的 `store.py` 把一切都摊成"一个键一个值"（`fact(key,value)`），
到了银行场景就不够了：一笔材料里有**多个主体**（借款人 / 保证人 / 法定代表人），
同一个键（如"名称"）分属不同主体；信息要能**按批次回溯**；被改过的值要能查出**改之前是什么**；
每一次动手都要留下**能重算的审计链**。v2 基础设计中的 15 张表就是为这些而建的（`27 §3.3`）。

本层的三条硬规矩（都来自口径，不是实现偏好）：

1. **只认 v2**。库不存在 → 建一个空的 v2 库；库是 v1 → **拒绝并告诉你跑 `db-migrate`**
   （绝不偷偷改一个银行库）。
2. **`fact` 行只追加**。同一个（主体，键）来了不同的值 → **先记冲突，不覆盖**；
   真要覆盖只能走 `db-merge` 的勾选，那时新行写入、旧行标 `superseded_by`，历史不丢（`17 §3.8`）。
3. **`source_kind` 只能三种**：`source`（源文件来的）/ `computed`（算出来的，须带 `formula`+`inputs`）
   / `user`（你确认的）。这就是红线二"不无中生有"的落点（`27 §7.1`）。

事件链公式照抄 `06 §2.2`，默认 SM3（D-5）：

    hash(n) = H( prev_hash ‖ event_type ‖ payload_hash ‖ occurred_at ‖ actor_id ‖ run_id ‖ entity_id ‖ seq )

⚠️ **诚实边界**：链能发现"记录被改过"，**不承诺防住有管理员权限的人重做整条记录**（`23 §3.1`）。
"""
from __future__ import annotations

import hashlib
import json
import re
import shutil
import sqlite3
import time
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Iterable, Sequence

from . import schema_v2 as S

SCHEMA_VERSION_V2 = S.SCHEMA_VERSION_V2

BATCH_RE = re.compile(r"^\d{8}-\d{2}$")

#: `fact.source_kind` 的封闭集合（红线二）
SOURCE_KINDS = ("source", "computed", "user")
#: `fact.status`（27 §3.3）
FACT_STATUS = ("ok", "missing", "suspect")
#: `fill_op.status`（27 §3.3；含 `running` 才能实现 TA6"无 running 残留"）
FILL_STATUS = ("running", "ok", "missing", "blocked", "skipped", "aborted", "rejected")
#: `review_queue.kind`
REVIEW_KINDS = ("field", "role", "conflict")
#: `case_role.role`（X-3：只做粗粒度，集合封闭）
CASE_ROLES = ("借款人", "保证人", "法定代表人")

#: 值就是主体名的字段（与 `schema_v2.SUBJECT_NAME_KEYS` 同源）
SUBJECT_NAME_KEYS = tuple(S.SUBJECT_NAME_KEYS)
#: 个人保证人姓名的值也是主体名，但主体类型是自然人
PERSONAL_NAME_KEYS = ("个人保证人姓名",)
#: 值**一定是自然人**的字段：法定代表人在法律上只能是自然人，个人保证人同理。
#: 这是"行内写清"的明确信号，不是从名字猜性别——`27 §3.4.2` 要求不许从名字猜。
NATURAL_PERSON_KEYS = ("借款人法定代表人", "保证人法定代表人", "个人保证人姓名")
#: 键前缀 → 该事实属于扮演这个角色的主体
ROLE_PREFIXES = (("借款人", "借款人"), ("保证人", "保证人"))


class SchemaError(RuntimeError):
    """库的版本不对，或者 DDL 与设计文档对不上。"""


def now_utc() -> str:
    """`occurred_at`：UTC 时刻 + 明确时区（`06 §2.2`，G-B8）。"""
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def digest(data: str, algo: str = "sm3") -> str:
    """国密 SM3 优先（D-5），本机不支持时退到 SHA-256 并让哈希长度自证。"""
    try:
        return hashlib.new(algo, data.encode("utf-8")).hexdigest()
    except Exception:  # noqa: BLE001 - 算法可配置，退到 sha256
        return hashlib.sha256(data.encode("utf-8")).hexdigest()


def sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as fh:
        for block in iter(lambda: fh.read(1 << 20), b""):
            h.update(block)
    return h.hexdigest()


def schema_version(con: sqlite3.Connection) -> int:
    return int(con.execute("PRAGMA user_version").fetchone()[0] or 0)


def require_v2(con: sqlite3.Connection) -> None:
    v = schema_version(con)
    if v != SCHEMA_VERSION_V2:
        raise SchemaError(
            "本命令只在 SCHEMA v2 上工作（当前 user_version=%d）。"
            "先跑 `office.py db-migrate --db <库>`。" % v)


RULE_PROTECTION_SQL = """
CREATE TABLE IF NOT EXISTS template_rule_disabled (
    template_id INTEGER NOT NULL, field TEXT NOT NULL,
    disabled_at TEXT NOT NULL, reason TEXT NOT NULL,
    PRIMARY KEY (template_id, field)
);
CREATE TRIGGER IF NOT EXISTS template_rule_no_delete
BEFORE DELETE ON template_rule
BEGIN
    SELECT RAISE(ABORT, 'template_rule 不允许按编号删除；请按模板和字段停用');
END;
"""

# Fill-card answers are workspace data. Keep the v2 base schema unchanged:
# existing databases acquire this small extension when opened, and new ones
# receive it immediately after the base tables are created.
FILL_DECISION_SQL = """
CREATE TABLE IF NOT EXISTS fill_decision (
    fingerprint TEXT PRIMARY KEY,
    batch_no TEXT NOT NULL,
    template_sha256 TEXT NOT NULL,
    field TEXT NOT NULL,
    action TEXT NOT NULL CHECK (action IN ('accept','blank','new','use')),
    value TEXT,
    decided_at TEXT NOT NULL,
    run_id TEXT NOT NULL
);
CREATE INDEX IF NOT EXISTS ix_fill_decision_batch ON fill_decision(batch_no);
"""


def open_db(db_path: str | Path, *, create: bool = True) -> sqlite3.Connection:
    """打开（必要时新建）一个 v2 库。**v1 库一律拒绝，绝不偷偷迁移。**"""
    p = Path(db_path)
    if not p.exists():
        if not create:
            raise SchemaError("数据库不存在：%s" % p)
        p.parent.mkdir(parents=True, exist_ok=True)
        con = sqlite3.connect(str(p))
        con.row_factory = sqlite3.Row
        con.executescript(S.DDL_V2)
        con.executescript(FILL_DECISION_SQL)
        con.execute("PRAGMA user_version = %d" % SCHEMA_VERSION_V2)
        con.commit()
        return con
    con = sqlite3.connect(str(p))
    con.row_factory = sqlite3.Row
    require_v2(con)
    con.executescript(RULE_PROTECTION_SQL)
    con.executescript(FILL_DECISION_SQL)
    con.commit()
    return con


# --------------------------------------------------------------------------
# 批次号：**唯一编号规则** `YYYYMMDD-NN`（26 §2.3）
# --------------------------------------------------------------------------
def valid_batch(value: str | None) -> bool:
    return bool(value) and bool(BATCH_RE.match(str(value)))


def allocate_batch(con: sqlite3.Connection, day: str | None = None,
                   label: str | None = None) -> str:
    """当天分配一个新的批次号；`label` 是人看的批次名，**不进批次号**（26 §2.3）。"""
    day = day or time.strftime("%Y%m%d")
    if not re.match(r"^\d{8}$", day):
        raise SchemaError("批次日必须是 YYYYMMDD：%r" % day)
    used = {r[0] for r in con.execute(
        "SELECT batch_no FROM source WHERE batch_no IS NOT NULL")}
    used |= {r[0] for r in con.execute(
        "SELECT batch_no FROM fact WHERE batch_no IS NOT NULL")}
    used |= {r[0] for r in con.execute(
        "SELECT batch_no FROM event WHERE batch_no IS NOT NULL")}
    n = 0
    while True:
        n += 1
        cand = "%s-%02d" % (day, n)
        if cand not in used:
            return cand


def resolve_batch(con: sqlite3.Connection, entity_id: int | None = None,
                  explicit: str | None = None) -> str:
    """给"补录/勾选覆盖"这类不属于某次新交付的动作找批次锚点。

    顺序：① 显式指定 ② 该主体最新一条 `fact` 的批次 ③ 库里最新的 `source` 批次
    ④ 当天分配一个新批次号。**不发明新的编号规则。**
    """
    if explicit:
        if not valid_batch(explicit):
            raise SchemaError("批次号必须是 YYYYMMDD-NN：%r" % explicit)
        return explicit
    if entity_id is not None:
        r = con.execute("SELECT batch_no FROM fact WHERE entity_id=? AND batch_no IS NOT NULL"
                        " ORDER BY id DESC LIMIT 1", (entity_id,)).fetchone()
        if r:
            return r[0]
    r = con.execute("SELECT batch_no FROM source WHERE batch_no IS NOT NULL"
                    " ORDER BY id DESC LIMIT 1").fetchone()
    if r:
        return r[0]
    return allocate_batch(con)


# --------------------------------------------------------------------------
# 审计事件链
# --------------------------------------------------------------------------
def write_event(con: sqlite3.Connection, *, batch_no: str, event_type: str, actor_id: str,
                entity_id: int | None = None, target: str | None = None,
                payload: dict | None = None, value_text: str | None = None,
                run_id: str | None = None, case_id: int | None = None,
                algo: str = "sm3") -> int:
    """按 `06 §2.2` 的公式追加一条审计事件，并返回它的 `id`。

    `previous_hash` 取该批次内 `seq` 最大的那一条的 `hash`；链条**按批次**分立。
    """
    payload = payload or {}
    payload_json = json.dumps(payload, ensure_ascii=False, sort_keys=True, default=str)
    payload_hash = digest(payload_json, algo)
    row = con.execute("SELECT seq, hash FROM event WHERE batch_no=? ORDER BY seq DESC LIMIT 1",
                      (batch_no,)).fetchone()
    seq = (row[0] + 1) if row else 1
    prev_hash = row[1] if row else None
    occurred_at = now_utc()
    material = "\u2016".join([
        prev_hash or "", event_type, payload_hash, occurred_at,
        actor_id or "", run_id or "", str(entity_id if entity_id is not None else ""), str(seq)])
    cur = con.execute(
        "INSERT INTO event(batch_no,run_id,seq,occurred_at,actor_id,actor_name,event_type,"
        "case_id,target,entity_id,payload_json,payload_hash,value_text,prev_hash,hash)"
        " VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
        (batch_no, run_id, seq, occurred_at, actor_id, None, event_type, case_id, target,
         entity_id, payload_json, payload_hash, value_text, prev_hash, digest(material, algo)))
    return int(cur.lastrowid)


def verify_chain(con: sqlite3.Connection, algo: str = "sm3") -> dict:
    """独立重算整条链（TA3 / TC3）。**只报告，不改任何东西。**"""
    prev: dict[str, str | None] = {}
    checked = 0
    for r in con.execute("SELECT batch_no,seq,event_type,payload_hash,occurred_at,actor_id,"
                         "run_id,entity_id,prev_hash,hash FROM event ORDER BY batch_no,seq"):
        bno, seq, etype, ph, at, actor, run_id, eid, stored_prev, stored_hash = r
        want_prev = prev.get(bno)
        if (stored_prev or None) != (want_prev or None):
            return {"ok": False, "checked": checked,
                    "problem": "批次 %s 第 %s 条的 prev_hash 对不上" % (bno, seq)}
        material = "\u2016".join([stored_prev or "", etype, ph, at, actor or "",
                                  run_id or "", str(eid if eid is not None else ""), str(seq)])
        if digest(material, algo) != stored_hash:
            return {"ok": False, "checked": checked,
                    "problem": "批次 %s 第 %s 条的 hash 算不回来（被改过？）" % (bno, seq)}
        prev[bno] = stored_hash
        checked += 1
    return {"ok": True, "checked": checked, "batches": len(prev)}


# --------------------------------------------------------------------------
# 键 → 主体的归属规则（**不许猜名字**，27 §3.4.2）
# --------------------------------------------------------------------------
def entity_type_for_name(name: str, *, personal: bool = False) -> str:
    return "natural_person" if personal else "legal_person"


def role_prefix_of(key: str) -> str | None:
    for prefix, role in ROLE_PREFIXES:
        if str(key).startswith(prefix):
            return role
    return None


class StoreV2:
    """v2 库的显式包装。所有写操作都在这里，方便一处看清"谁写了什么"。"""

    def __init__(self, path: str | Path, *, actor: str = "user", create: bool = True):
        self.path = Path(path)
        self.actor = actor
        self.conn = open_db(self.path, create=create)
        self.conn.execute("PRAGMA foreign_keys = ON")

    # ---- 生命周期 -----------------------------------------------------
    def close(self) -> None:
        try:
            self.conn.commit()
        finally:
            self.conn.close()

    def __enter__(self) -> "StoreV2":
        return self

    def __exit__(self, *exc) -> None:
        self.close()

    # ---- 事件 ---------------------------------------------------------
    def event(self, event_type: str, *, batch_no: str, entity_id: int | None = None,
              target: str | None = None, payload: dict | None = None,
              value_text: str | None = None, run_id: str | None = None,
              case_id: int | None = None, actor: str | None = None) -> int:
        eid = write_event(self.conn, batch_no=batch_no, event_type=event_type,
                          actor_id=actor or self.actor, entity_id=entity_id, target=target,
                          payload=payload, value_text=value_text, run_id=run_id,
                          case_id=case_id)
        self.conn.commit()
        return eid

    # ---- source -------------------------------------------------------
    def register_source(self, path: Path, batch_no: str, *, kind: str | None = None,
                        copy_root: Path | None = None, parse_status: str = "ok",
                        parse_note: str | None = None,
                        batch_label: str | None = None) -> tuple[int, bool]:
        """登记一份源文件；**可选存整份副本**到 `<copy_root>\\<批次号>\\`（D-1 第 2 轮）。

        返回 `(source_id, already)`。同一 (路径, 哈希) 只登记一次。
        """
        digest_ = sha256_file(path)
        row = self.conn.execute("SELECT id FROM source WHERE path=? AND sha256=?",
                                (str(path), digest_)).fetchone()
        if row:
            return int(row["id"]), True
        copy_path = None
        if copy_root is not None:
            dest = Path(copy_root) / batch_no / path.name
            dest.parent.mkdir(parents=True, exist_ok=True)
            if not dest.exists():
                shutil.copy2(path, dest)
            copy_path = str(dest)
        cur = self.conn.execute(
            "INSERT INTO source(batch_no,batch_label,path,copy_path,name,sha256,bytes,kind,"
            "parse_status,parse_note,ingested_at) VALUES(?,?,?,?,?,?,?,?,?,?,?)",
            (batch_no, batch_label, str(path), copy_path, path.name, digest_,
             path.stat().st_size, kind or path.suffix.lstrip("."), parse_status,
             parse_note, now_utc()))
        self.conn.commit()
        sid = int(cur.lastrowid)
        self.event("source_registered", batch_no=batch_no, target=path.name,
                   payload={"source_id": sid, "path": str(path), "sha256": digest_,
                            "copy_path": copy_path, "parse_status": parse_status,
                            "parse_note": parse_note, "bytes": path.stat().st_size})
        return sid, False

    def register_virtual_source(self, name: str, batch_no: str, *, kind: str = "profile",
                                note: str | None = None, sha256: str = "",
                                bytes_: int = 0) -> int:
        """登记一个**没有实体文件**的来源（如"人工确认"）——出处照样可回溯。"""
        row = self.conn.execute("SELECT id FROM source WHERE path=? AND sha256=?",
                                (name, sha256)).fetchone()
        if row:
            return int(row["id"])
        cur = self.conn.execute(
            "INSERT INTO source(batch_no,batch_label,path,copy_path,name,sha256,bytes,kind,"
            "parse_status,parse_note,ingested_at) VALUES(?,?,?,?,?,?,?,?,?,?,?)",
            (batch_no, None, name, None, name, sha256, bytes_, kind, "ok", note, now_utc()))
        self.conn.commit()
        return int(cur.lastrowid)

    def sources(self, batch_no: str | None = None) -> list[dict]:
        sql = "SELECT * FROM source"
        args: tuple = ()
        if batch_no:
            sql += " WHERE batch_no=?"
            args = (batch_no,)
        return [dict(r) for r in self.conn.execute(sql + " ORDER BY id", args)]

    # ---- entity -------------------------------------------------------
    def find_entity_by_name(self, name: str) -> int | None:
        r = self.conn.execute(
            "SELECT entity_id FROM entity_name WHERE name=? ORDER BY"
            " CASE name_kind WHEN '现用名' THEN 0 ELSE 1 END, id LIMIT 1",
            (str(name).strip(),)).fetchone()
        return int(r[0]) if r else None

    def find_entity_by_key(self, kind: str, value: str) -> int | None:
        col = {"uscc": "uscc", "bank_no": "bank_no", "id_card": "id_card"}[kind]
        r = self.conn.execute('SELECT id FROM entity WHERE "%s"=?' % col,
                              (str(value).strip(),)).fetchone()
        return int(r[0]) if r else None

    def ensure_entity(self, name: str, *, entity_type: str | None = None,
                      personal: bool = False) -> tuple[int, bool]:
        """按名字找主体；找不到就建档。识别码缺失仅记为提示。"""
        name = str(name).strip()
        if not name:
            raise SchemaError("主体名不能为空")
        eid = self.find_entity_by_name(name)
        if eid is not None:
            return eid, False
        et = entity_type or entity_type_for_name(name, personal=personal)
        ts = now_utc()
        cur = self.conn.execute(
            "INSERT INTO entity(entity_type,identity_incomplete,note,created_at,updated_at)"
            " VALUES(?,1,?,?,?)", (et, "由源文件中的主体名自动建档；识别码可选", ts, ts))
        eid = int(cur.lastrowid)
        self.conn.execute("INSERT INTO entity_name(entity_id,name,name_kind,created_at)"
                          " VALUES(?,?,'现用名',?)", (eid, name, ts))
        self.conn.commit()
        return eid, True

    def entity_label(self, entity_id: int | None) -> str:
        if entity_id is None:
            return "（未归属）"
        r = self.conn.execute(
            "SELECT name FROM entity_name WHERE entity_id=? ORDER BY"
            " CASE name_kind WHEN '现用名' THEN 0 ELSE 1 END, id LIMIT 1",
            (entity_id,)).fetchone()
        return r[0] if r else "主体#%s" % entity_id

    def entity_keys(self, entity_id: int) -> dict:
        r = self.conn.execute("SELECT * FROM entity WHERE id=?", (entity_id,)).fetchone()
        return dict(r) if r else {}

    def entities(self) -> list[dict]:
        out = []
        for r in self.conn.execute("SELECT * FROM entity ORDER BY id"):
            d = dict(r)
            d["name"] = self.entity_label(d["id"])
            d["roles"] = self.roles_of(d["id"])
            out.append(d)
        return out

    def ensure_case(self, entity_id: int, batch_no: str | None = None,
                    *, facility_no: str | None = None) -> int:
        """一个主体一个"空壳案子"（额度层，26 §2.2）；已有就复用最新的那个。"""
        r = self.conn.execute('SELECT id FROM "case" WHERE entity_id=? ORDER BY id DESC LIMIT 1',
                              (entity_id,)).fetchone()
        if r:
            return int(r[0])
        cur = self.conn.execute(
            'INSERT INTO "case"(batch_no,entity_id,facility_no,created_at) VALUES(?,?,?,?)',
            (batch_no, entity_id, facility_no, now_utc()))
        cid = int(cur.lastrowid)
        self.conn.commit()
        self.event("case_created", batch_no=batch_no or resolve_batch(self.conn, entity_id),
                   entity_id=entity_id, case_id=cid, target=self.entity_label(entity_id),
                   payload={"case_id": cid, "entity_id": entity_id, "facility_no": facility_no})
        return cid

    def latest_case(self, entity_id: int) -> int | None:
        r = self.conn.execute('SELECT id FROM "case" WHERE entity_id=? ORDER BY id DESC LIMIT 1',
                              (entity_id,)).fetchone()
        return int(r[0]) if r else None

    # ---- case_role ----------------------------------------------------
    def set_role(self, entity_id: int, role: str, *, case_id: int | None = None,
                 evidence: str | None = None, decided_by: str = "human",
                 source_id: int | None = None, batch_no: str | None = None) -> int:
        """记录主体在案子里的角色。`role` 必须是三选一（X-3 粗粒度）。"""
        if role not in CASE_ROLES:
            raise SchemaError("角色只能是 %s（X-3：只做粗粒度，集合封闭）：%r"
                              % ("/".join(CASE_ROLES), role))
        cid = case_id or self.latest_case(entity_id) or self.ensure_case(entity_id, batch_no)
        bno = batch_no or resolve_batch(self.conn, entity_id)
        row = self.conn.execute("SELECT id FROM case_role WHERE entity_id=? AND case_id=? AND role=?",
                                (entity_id, cid, role)).fetchone()
        if row:
            self.conn.execute("UPDATE case_role SET evidence=?, decided_by=?, source_id=?"
                              " WHERE id=?", (evidence, decided_by, source_id, int(row[0])))
            self.conn.commit()
            return int(row[0])
        cur = self.conn.execute(
            "INSERT INTO case_role(entity_id,case_id,role,evidence,decided_by,source_id,created_at)"
            " VALUES(?,?,?,?,?,?,?)",
            (entity_id, cid, role, evidence, decided_by, source_id, now_utc()))
        self.conn.commit()
        self.event("case_role_decided", batch_no=bno, entity_id=entity_id, case_id=cid,
                   target=self.entity_label(entity_id),
                   payload={"entity_id": entity_id, "case_id": cid, "role": role,
                            "decided_by": decided_by, "evidence": evidence})
        return int(cur.lastrowid)

    def roles_of(self, entity_id: int) -> list[str]:
        return [r[0] for r in self.conn.execute(
            "SELECT DISTINCT role FROM case_role WHERE entity_id=? ORDER BY role", (entity_id,))]

    def roles(self) -> list[dict]:
        out = []
        for r in self.conn.execute("SELECT * FROM case_role ORDER BY entity_id, role"):
            d = dict(r)
            d["entity"] = self.entity_label(d["entity_id"])
            out.append(d)
        return out

    def entity_for_role(self, role: str, *, case_id: int | None = None) -> int | None:
        if case_id is not None:
            r = self.conn.execute("SELECT entity_id FROM case_role WHERE role=? AND case_id=?"
                                  " ORDER BY id LIMIT 1", (role, case_id)).fetchone()
            if r:
                return int(r[0])
        r = self.conn.execute("SELECT entity_id FROM case_role WHERE role=? ORDER BY id LIMIT 1",
                              (role,)).fetchone()
        return int(r[0]) if r else None

    # ---- fact ---------------------------------------------------------
    def _current_row(self, key: str, entity_id: int | None) -> sqlite3.Row | None:
        return self.current_row(key, entity_id)

    def current_row(self, key: str, entity_id: int | None) -> sqlite3.Row | None:
        """这个（主体，键）**当前有效**的那一行；没有则 None。"""
        if entity_id is None:
            return self.conn.execute(
                "SELECT * FROM fact WHERE key=? AND entity_id IS NULL AND superseded_by IS NULL"
                " ORDER BY id DESC LIMIT 1", (key,)).fetchone()
        return self.conn.execute(
            "SELECT * FROM fact WHERE key=? AND entity_id=? AND superseded_by IS NULL"
            " ORDER BY id DESC LIMIT 1", (key, entity_id)).fetchone()

    def put_fact(self, batch_no: str, key: str, value: Any, *, entity_id: int | None = None,
                 case_id: int | None = None, source_id: int | None = None,
                 source_kind: str = "source", provenance: str | None = None,
                 unit: str | None = None, value_num: float | None = None,
                 confidence: float | None = None, status: str = "ok",
                 note: str | None = None, agreement_no: str | None = None,
                 formula: str | None = None, inputs: Any = None,
                 effective_from: str | None = None,
                 on_conflict: str = "record") -> dict:
        """写入一条键值对。

        `on_conflict`：
          * `record`（默认）—— 已有不同的值 → **只记冲突，不覆盖**（要覆盖请走 `db-merge`）
          * `supersede`     —— 写入新行并把旧行标 `superseded_by`（**只有 `db-merge` 的勾选该用**）
          * `ignore`        —— 已有值就什么都不做

        返回值说清楚到底发生了什么，调用方**不许把它当成"写成功了"就完事**。
        """
        if source_kind not in SOURCE_KINDS:
            raise SchemaError("source_kind 只能是 %s（红线二）：%r"
                              % ("/".join(SOURCE_KINDS), source_kind))
        if source_kind == "computed" and not (formula and inputs is not None):
            raise SchemaError("source_kind='computed' 必须同时给 formula 与 inputs（27 §7.1 红线二）")
        if status not in FACT_STATUS:
            raise SchemaError("fact.status 只能是 %s：%r" % ("/".join(FACT_STATUS), status))
        text = None if value is None else str(value)
        old = self.conn.execute("SELECT * FROM fact WHERE batch_no=? AND key=? AND entity_id IS ? AND superseded_by IS NULL ORDER BY id DESC LIMIT 1", (batch_no, key, entity_id)).fetchone()

        # 库里那一行是**空的**（status='missing' 或值为空），这次有了值：
        # 这是"把空白补上"，**不是覆盖**——补空白不可能丢信息，所以直接写，
        # 也不该记成冲突（否则每一次"原来缺、后来补上了"都会变成一条待裁定的冲突）。
        if old is not None and not (old["value"] or "").strip() and (text or "").strip():
            on_conflict = "supersede"

        if old is not None:
            same = (old["value"] or "") == (text or "")
            if same:
                return {"fact_id": int(old["id"]), "action": "unchanged",
                        "key": key, "entity_id": entity_id}
            if on_conflict == "ignore":
                return {"fact_id": int(old["id"]), "action": "kept_existing",
                        "key": key, "entity_id": entity_id, "existing": old["value"]}
            if on_conflict == "record":
                cid = self.conn.execute(
                    "INSERT INTO fact_conflict(batch_no,key,entity_id,agreement_no,existing,"
                    "incoming,existing_src,incoming_src,norm_rule,detected_at)"
                    " VALUES(?,?,?,?,?,?,?,?,?,?)",
                    (batch_no, key, entity_id, agreement_no, old["value"], text,
                     old["source_id"], source_id, _norm_rule_for(key), now_utc())).lastrowid
                self.conn.commit()
                self.event("fact_conflict_detected", batch_no=batch_no, entity_id=entity_id,
                           target=key, payload={"key": key, "existing": old["value"],
                                                "incoming": text, "conflict_id": int(cid),
                                                "entity_id": entity_id},
                           value_text=text)
                return {"fact_id": int(old["id"]), "action": "conflict",
                        "conflict_id": int(cid), "key": key, "entity_id": entity_id,
                        "existing": old["value"], "incoming": text}

        cur = self.conn.execute(
            "INSERT INTO fact(batch_no,source_id,entity_id,case_id,agreement_no,key,value,unit,"
            "value_num,confidence,source_kind,provenance,formula,inputs,status,effective_from,"
            "note,created_at) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
            (batch_no, source_id, entity_id, case_id, agreement_no, key, text, unit, value_num,
             confidence, source_kind, provenance, formula,
             None if inputs is None else json.dumps(inputs, ensure_ascii=False, default=str),
             status, effective_from, note, now_utc()))
        fid = int(cur.lastrowid)
        superseded = None
        if old is not None and on_conflict == "supersede":
            self.conn.execute("UPDATE fact SET superseded_by=? WHERE id=?", (fid, int(old["id"])))
            superseded = int(old["id"])
        self.conn.commit()
        self.event("fact_extracted", batch_no=batch_no, entity_id=entity_id, target=key,
                   payload={"fact_id": fid, "key": key, "entity_id": entity_id,
                            "source_kind": source_kind, "provenance": provenance,
                            "source_id": source_id, "status": status,
                            "superseded": superseded},
                   value_text=text)
        if superseded is not None:
            self.event("fact_superseded", batch_no=batch_no, entity_id=entity_id, target=key,
                       payload={"key": key, "entity_id": entity_id, "old_fact_id": superseded,
                                "new_fact_id": fid,
                                "old_value": old["value"], "new_value": text,
                                "why": "勾选覆盖（db-merge）"},
                       value_text=text)
        return {"fact_id": fid, "action": "superseded" if superseded else "inserted",
                "key": key, "entity_id": entity_id, "superseded_fact_id": superseded}

    def discard_fact(self, batch_no: str, key: str, value: Any, *, entity_id: int | None = None,
                     reason: str = "源侧复核：丢弃") -> int:
        """源侧"丢弃"：**不写库，但必须留痕**（`fact_discarded`，26 §3.1）。

        ⚠️ 不能拿 `fact_superseded` 顶替——那个要求有"新值"，硬塞就是伪造一条"更正"。
        """
        return self.event("fact_discarded", batch_no=batch_no, entity_id=entity_id, target=key,
                          payload={"key": key, "entity_id": entity_id, "value": value,
                                   "reason": reason},
                          value_text=None if value is None else str(value))

    def current_facts(self, *, entity_id: int | None = None,
                      include_missing: bool = True) -> dict[str, dict]:
        """每个键**当前的**值（已 `superseded_by` 的旧行不算）。

        键用 `键` 或 `主体#id|键`：同一键属于多个主体时自动带上主体，避免张冠李戴。
        """
        sql = ("SELECT f.*, e.entity_type FROM fact f LEFT JOIN entity e ON e.id=f.entity_id"
               " WHERE f.superseded_by IS NULL")
        args: tuple = ()
        if entity_id is not None:
            sql += " AND f.entity_id=?"
            args = (entity_id,)
        rows = [dict(r) for r in self.conn.execute(sql + " ORDER BY f.id", args)]
        counts: dict[str, int] = {}
        for r in rows:
            counts[r["key"]] = counts.get(r["key"], 0) + 1
        out: dict[str, dict] = {}
        for r in rows:
            label = r["key"] if counts[r["key"]] == 1 else "%s#%s|%s" % (
                self.entity_label(r["entity_id"]), r["entity_id"], r["key"])
            r["display_key"] = label
            if not include_missing and not (r["value"] or "").strip():
                continue
            out[label] = r
        return out

    def facts_by_key(self) -> dict[str, dict]:
        """扁平视图：`fact.key` → 那一行。

        🔴 **同一键属于多个主体时，这里返回的是"歧义占位"，绝不挑一条给你**
        （T-25，2026-09-21 修）。

        为什么必须这么改（**实测过的真事故**）：原来这里是"同一键取最新一行"。
        甲公司先入库、乙公司后入库，两家都有 `联系电话` —— 于是**后录入的乙公司电话赢了**，
        表单上那个不带主体限定的「联系电话：」格子**静默填了乙公司的电话**，
        `decision='auto'`、没有任何提示，而且因为 `map_role('联系电话')` 认不出角色，
        **连红线一的跨主体核对都跳过了**。

        歧义占位的形状：::

            {"key": k, "value": None, "entity_id": None, "_ambiguous": True,
             "_candidates": [{"fact_id", "entity_id", "entity_name", "value"}, ...]}

        ⚠️ 老代码里 `_ambiguous` 只是**标了但没人读**（留了位没接线）。
        现在它带着 `_candidates` 一起返回，调用方**必须**决定怎么办——
        `build_fill_plan` 会把它变成"**要你点头：请指明用谁的值**"。
        """
        rows = [dict(r) for r in self.conn.execute(
            "SELECT * FROM fact WHERE superseded_by IS NULL AND (? IS NULL OR batch_no=?) ORDER BY id", (getattr(self, "batch_scope", None), getattr(self, "batch_scope", None)))]
        grouped: dict[str, list[dict]] = {}
        for r in rows:
            grouped.setdefault(r["key"], []).append(r)
        out: dict[str, dict] = {}
        for key, group in grouped.items():
            if len(group) == 1:
                out[key] = group[0]
                continue
            # 多行：只有**同一个主体**的重复行才算"没歧义"（取最新）
            eids = {r["entity_id"] for r in group}
            if len(eids) == 1:
                out[key] = group[-1]
                continue
            out[key] = {
                "key": key, "value": None, "entity_id": None, "_ambiguous": True,
                "_candidates": [{"fact_id": r["id"], "entity_id": r["entity_id"],
                                 "entity_name": self.entity_label(r["entity_id"]),
                                 "value": r["value"], "provenance": r["provenance"]}
                                for r in group],
            }
        return out

    def current_rows(self, key: str) -> list[dict]:
        """这个键**当前有效**的所有行（按主体分开，不去重、不挑）。"""
        return [dict(r) for r in self.conn.execute(
            "SELECT * FROM fact WHERE key=? AND superseded_by IS NULL AND (? IS NULL OR batch_no=?) ORDER BY id", (key, getattr(self, "batch_scope", None), getattr(self, "batch_scope", None)))]

    def entities_for_role(self, role: str, *, entity_type: str | None = None) -> list[int]:
        """Return every entity with ``role``; never silently pick the first one."""
        sql = ("SELECT DISTINCT cr.entity_id FROM case_role cr "
               "JOIN entity e ON e.id=cr.entity_id WHERE cr.role=?")
        args: list[Any] = [role]
        if entity_type is not None:
            sql += " AND e.entity_type=?"
            args.append(entity_type)
        return [int(r[0]) for r in self.conn.execute(sql + " ORDER BY cr.entity_id", args)]

    def is_company_evidenced(self, entity_id: int) -> bool:
        """Whether stored evidence, rather than a name or default type, identifies a company."""
        entity = self.conn.execute("SELECT entity_type,uscc,bank_no FROM entity WHERE id=?",
                                   (entity_id,)).fetchone()
        if entity is None or entity["entity_type"] == "natural_person":
            return False
        if entity["uscc"]:
            return True
        corporate_facts = ("统一社会信用代码", "借款人法定代表人", "保证人法定代表人", "法定代表人")
        hit = self.conn.execute(
            "SELECT 1 FROM fact WHERE entity_id=? AND key IN (%s) AND value IS NOT NULL "
            "AND TRIM(CAST(value AS TEXT))<>'' AND superseded_by IS NULL LIMIT 1"
            % ",".join("?" for _ in corporate_facts),
            (entity_id, *corporate_facts)).fetchone()
        return hit is not None

    def facts_for_subject(self, subject_eid: int | None, *,
                          scope_entity_ids: Iterable[int] | None = None) -> dict[str, dict]:
        """**按本次主体取值**的视图（T-25 的核心）。

        规则（与 `27 §7.1` 红线一同源，**不猜**）：

        1. 先按键名声称的角色，或本模板明确的主体范围过滤；
        2. 再在过滤后的候选里取唯一值，否则留歧义给使用者；
        3. 没有主体归属的全局事实可作为最后回退。

        不能因为某个键恰好只有一条记录，就把保证人的电话借给借款人模板。
        """
        from .schema_v2 import map_role

        scope_given = scope_entity_ids is not None
        scope = {int(e) for e in (scope_entity_ids or [])}
        if subject_eid is not None:
            scope.add(int(subject_eid))
        out: dict[str, dict] = {}
        for key, rows in ((k, self.current_rows(k)) for k in
                          {r["key"] for r in self.conn.execute(
                              "SELECT DISTINCT key FROM fact WHERE superseded_by IS NULL")}):
            if not rows:
                continue
            claimed = map_role(key)
            global_rows = [r for r in rows if r["entity_id"] is None]
            if claimed:
                role_eids = set(self.entities_for_role(claimed))
                role_rows = [r for r in rows if r["entity_id"] is not None
                             and int(r["entity_id"]) in role_eids]
                # A template's same-role scope is narrower than the whole
                # case.  For example, a company-guarantor template must not
                # reintroduce a personal guarantor merely because this field
                # also says "保证人".  A borrower template mentioning a
                # guarantor field has no such intersection and keeps it.
                if scope & role_eids:
                    candidates = [r for r in role_rows if int(r["entity_id"]) in scope]
                elif role_rows:
                    candidates = role_rows
                elif claimed == "法定代表人":
                    # The source's composite keys already carry the company
                    # relation (借款人法定代表人 / 保证人法定代表人).  Existing
                    # data may not have separately stored a person role, so
                    # retain its explicit candidates rather than discard it.
                    company_role = next((role for role in ('借款人', '保证人')
                                         if key.startswith(role)), None)
                    company_ids = set(self.entities_for_role(company_role)) if company_role else set()
                    matching_scope = scope & company_ids
                    candidates = [r for r in rows if r["entity_id"] is not None
                                  and (not matching_scope or int(r["entity_id"]) in matching_scope)]
                else:
                    candidates = []
            elif scope:
                candidates = [r for r in rows if r["entity_id"] is not None
                              and int(r["entity_id"]) in scope]
            elif scope_given:
                # An explicit template declaration had no matching entity.
                # Do not fall back to an unrelated single value.
                candidates = []
            else:
                # No template subject means facts tied to a person/company remain
                # ambiguous.  A truly global fact is still usable below.
                candidates = [r for r in rows if r["entity_id"] is not None]
            unresolved_scope = (subject_eid is None and scope_given and bool(scope)
                                and (not claimed or bool(scope & set(self.entities_for_role(claimed)))))
            if len(candidates) == 1 and not unresolved_scope:
                out[key] = candidates[0]
                continue
            if not candidates and len(global_rows) == 1:
                out[key] = global_rows[0]
                continue
            ambiguous_rows = candidates or global_rows
            if not ambiguous_rows:
                # The key exists, but none of its values belongs to this template.
                # Deliberately omit it so the plan says the scoped source lacks it.
                continue
            out[key] = {
                "key": key, "value": None, "entity_id": None, "_ambiguous": True,
                "_candidates": [{"fact_id": r["id"], "entity_id": r["entity_id"],
                                  "entity_name": self.entity_label(r["entity_id"]),
                                  "value": r["value"], "provenance": r["provenance"]}
                                 for r in ambiguous_rows],
            }
        return out

    def role_entity(self, role: str) -> int | None:
        """扮演这个角色的主体（`case_role` 里查）。"""
        return self.entity_for_role(role)

    def unassigned_facts(self) -> list[dict]:
        """`entity_id` 为空的事实——红线一**核对不了**的那些，必须让人看见。"""
        return [dict(r) for r in self.conn.execute(
            "SELECT * FROM fact WHERE entity_id IS NULL AND superseded_by IS NULL ORDER BY id")]

    def conflicts(self, *, open_only: bool = True) -> list[dict]:
        sql = "SELECT * FROM fact_conflict"
        if open_only:
            sql += " WHERE resolved=0"
        return [dict(r) for r in self.conn.execute(sql + " ORDER BY id")]

    def resolve_conflict(self, conflict_id: int, resolution: str, *, decided_by: str = "user",
                         batch_no: str | None = None) -> None:
        if resolution not in ("took_existing", "took_incoming", "left_blank"):
            raise SchemaError("裁定只能是 took_existing / took_incoming / left_blank：%r" % resolution)
        self.conn.execute("UPDATE fact_conflict SET resolved=1, resolution=?, resolved_by=?,"
                          " resolved_at=? WHERE id=?",
                          (resolution, decided_by, now_utc(), conflict_id))
        self.conn.commit()

    # ---- template / template_rule -------------------------------------
    def register_template(self, path: Path, batch_no: str, *, doc_kind: str | None = None,
                          structure_ok: int | None = None,
                          structure_note: str | None = None) -> int:
        digest_ = sha256_file(path)
        r = self.conn.execute("SELECT id,sha256 FROM template WHERE path=?", (str(path),)).fetchone()
        if r:
            if r["sha256"] != digest_:
                self.conn.execute(
                    "INSERT OR IGNORE INTO template_rule_disabled"
                    "(template_id,field,disabled_at,reason)"
                    " SELECT template_id,field,?,? FROM template_rule WHERE template_id=?",
                    (now_utc(), "模板文件内容已变化，旧填写位置停用", int(r["id"])))
            self.conn.execute("UPDATE template SET sha256=?, registered_at=? WHERE id=?",
                              (digest_, now_utc(), int(r["id"])))
            self.conn.commit()
            return int(r["id"])
        cur = self.conn.execute(
            "INSERT INTO template(batch_no,path,name,sha256,doc_kind,structure_ok,structure_note,"
            "registered_at) VALUES(?,?,?,?,?,?,?,?)",
            (batch_no, str(path), path.name, digest_,
             doc_kind or path.suffix.lstrip(".").lower(), structure_ok, structure_note, now_utc()))
        self.conn.commit()
        return int(cur.lastrowid)

    def add_rule(self, template_id: int, field: str, label: str | None, target: dict,
                 *, is_required: int = 0, match_kind: str = "exact",
                 confidence: float | None = None, decided_by: str = "human",
                 batch_no: str | None = None) -> int:
        """一行 = 目标表单上的**一个格子**（`26 §2.7`：应填字段数 = 行数，完整率的分母）。"""
        if match_kind not in ("exact", "normalized"):
            raise SchemaError("match_kind 只能是 exact / normalized：%r" % match_kind)
        if decided_by not in ("human", "model", "auto"):
            raise SchemaError("decided_by 只能是 human / model / auto：%r" % decided_by)
        row = self.conn.execute("SELECT id FROM template_rule WHERE template_id=? AND field=?",
                                (template_id, field)).fetchone()
        if row:
            self.conn.execute("UPDATE template_rule SET label=?, target_json=?, is_required=?,"
                              " match_kind=?, confidence=?, decided_by=? WHERE id=?",
                              (label, json.dumps(target, ensure_ascii=False), is_required,
                               match_kind, confidence, decided_by, int(row[0])))
            self.conn.execute("DELETE FROM template_rule_disabled WHERE template_id=? AND field=?",
                              (template_id, field))
            self.conn.commit()
            return int(row[0])
        cur = self.conn.execute(
            "INSERT INTO template_rule(template_id,field,label,target_json,is_required,match_kind,"
            "confidence,decided_by,created_at) VALUES(?,?,?,?,?,?,?,?,?)",
            (template_id, field, label, json.dumps(target, ensure_ascii=False), is_required,
             match_kind, confidence, decided_by, now_utc()))
        self.conn.execute("DELETE FROM template_rule_disabled WHERE template_id=? AND field=?",
                          (template_id, field))
        self.conn.commit()
        # ⚠️ 规则本身**不写 event**：`06 §2.1` 的枚举里没有规则类事件
        #    （`rule_proposed`/`rule_approved` 四个已按 N-1 删除），不另造名字。
        return int(cur.lastrowid)

    def disable_rule(self, template_path: str | Path, field: str,
                     expected_target: dict, reason: str) -> int:
        """Disable by stable template+field and checked current position; retain history."""
        row = self.conn.execute(
            "SELECT r.id, r.target_json, t.sha256 FROM template_rule r"
            " JOIN template t ON t.id=r.template_id WHERE t.path=? AND r.field=?",
            (str(template_path), field)).fetchone()
        if not row or json.loads(row["target_json"]) != expected_target \
                or row["sha256"] != sha256_file(Path(template_path)):
            raise SchemaError("规则已变化或位置不符，请重新查看当前模板和规则，不能按旧编号清理")
        self.conn.execute(
            "INSERT INTO template_rule_disabled(template_id,field,disabled_at,reason)"
            " SELECT template_id,field,?,? FROM template_rule WHERE id=?"
            " ON CONFLICT(template_id,field) DO UPDATE SET disabled_at=excluded.disabled_at,"
            " reason=excluded.reason",
            (now_utc(), reason, row["id"]))
        self.conn.commit()
        return int(row["id"])

    def rules_for(self, template_path: str | Path) -> list[dict]:
        rows = self.conn.execute(
            "SELECT r.*, t.path AS template_path, t.name AS template_name, t.sha256 AS tpl_sha256"
            " FROM template_rule r JOIN template t ON t.id = r.template_id"
            " WHERE t.path=? AND NOT EXISTS (SELECT 1 FROM template_rule_disabled d"
            " WHERE d.template_id=r.template_id AND d.field=r.field) ORDER BY r.id",
            (str(template_path),)).fetchall()
        out = []
        for r in rows:
            d = dict(r)
            d["target"] = json.loads(d.pop("target_json"))
            out.append(d)
        return out

    # ---- fill_op（三种行：run / artifact / field） ---------------------
    def start_run(self, batch_no: str, run_id: str, *, preflight: dict | None = None) -> int:
        cur = self.conn.execute(
            "INSERT INTO fill_op(run_id,batch_no,kind,status,preflight_json,created_at)"
            " VALUES(?,?,'run','running',?,?)",
            (run_id, batch_no, None if preflight is None else json.dumps(
                preflight, ensure_ascii=False, default=str), now_utc()))
        self.conn.commit()
        return int(cur.lastrowid)

    def set_run_status(self, run_id: str, status: str, detail: str = "") -> None:
        if status not in FILL_STATUS:
            raise SchemaError("fill_op.status 只能是 %s：%r" % ("/".join(FILL_STATUS), status))
        self.conn.execute("UPDATE fill_op SET status=?, detail=? WHERE run_id=? AND kind='run'",
                          (status, detail, run_id))
        self.conn.commit()

    def record_field(self, run_id: str, batch_no: str, field: str, status: str, *,
                     entity_id: int | None = None, template_id: int | None = None,
                     template_rule_id: int | None = None, fact_id: int | None = None,
                     value: str | None = None, detail: str = "") -> int:
        if status not in FILL_STATUS:
            raise SchemaError("fill_op.status 只能是 %s：%r" % ("/".join(FILL_STATUS), status))
        cur = self.conn.execute(
            "INSERT INTO fill_op(run_id,batch_no,kind,entity_id,template_id,template_rule_id,"
            "fact_id,field,value,status,detail,created_at)"
            " VALUES(?,?,'field',?,?,?,?,?,?,?,?,?)",
            (run_id, batch_no, entity_id, template_id, template_rule_id, fact_id, field,
             value, status, detail, now_utc()))
        self.conn.commit()
        return int(cur.lastrowid)

    def record_artifact(self, run_id: str, batch_no: str, *, artifact_path: str,
                        artifact_sha256: str, template_id: int | None = None,
                        entity_id: int | None = None, opened_ok: bool | None = None,
                        status: str = "ok", detail: str = "") -> int:
        cur = self.conn.execute(
            "INSERT INTO fill_op(run_id,batch_no,kind,entity_id,template_id,artifact_path,"
            "artifact_sha256,opened_ok,status,detail,created_at)"
            " VALUES(?,?,'artifact',?,?,?,?,?,?,?,?)",
            (run_id, batch_no, entity_id, template_id, artifact_path, artifact_sha256,
             1 if opened_ok else (0 if opened_ok is not None else None), status, detail, now_utc()))
        self.conn.commit()
        return int(cur.lastrowid)

    def artifacts(self, *, run_id: str | None = None, batch_no: str | None = None) -> list[dict]:
        sql = ("SELECT a.*, e.entity_type, t.name AS template_name FROM fill_op a"
               " LEFT JOIN entity e ON e.id=a.entity_id"
               " LEFT JOIN template t ON t.id=a.template_id WHERE a.kind='artifact'")
        args: list = []
        if run_id:
            sql += " AND a.run_id=?"
            args.append(run_id)
        if batch_no:
            sql += " AND a.batch_no=?"
            args.append(batch_no)
        return [dict(r) for r in self.conn.execute(sql + " ORDER BY a.id", tuple(args))]

    def delivered_fields(self) -> dict[str, list[dict]]:
        """哪些字段**已经填进过产物**（`db-merge` 用它提醒"已交付的不受影响"）。"""
        out: dict[str, list[dict]] = {}
        for r in self.conn.execute(
                "SELECT f.field, f.value, f.created_at, a.artifact_path, a.artifact_sha256,"
                " a.signoff_decision, f.run_id FROM fill_op f"
                " LEFT JOIN fill_op a ON a.run_id=f.run_id AND a.kind='artifact'"
                " WHERE f.kind='field' AND f.status='ok' ORDER BY f.id"):
            out.setdefault(r["field"], []).append(dict(r))
        return out

    def signoff(self, run_id: str, *, decision: str, actor: str, scope: str = "batch",
                fingerprint: str | None = None) -> int:
        if decision not in ("approve", "reject"):
            raise SchemaError("签核只能是 approve / reject：%r" % decision)
        if scope not in ("batch", "document"):
            raise SchemaError("签核范围只能是 batch / document：%r" % scope)
        cur = self.conn.execute(
            "UPDATE fill_op SET signoff_scope=?, signoff_decision=?, signoff_actor=?,"
            " signoff_at=?, signoff_fingerprint=? WHERE run_id=? AND kind='artifact'",
            (scope, decision, actor, now_utc(), fingerprint, run_id))
        self.conn.commit()
        return int(cur.rowcount or 0)

    # ---- review_queue --------------------------------------------------
    def add_review(self, batch_no: str, kind: str, reason: str, *, run_id: str | None = None,
                   entity_id: int | None = None, case_id: int | None = None,
                   template_id: int | None = None, field: str | None = None,
                   label: str | None = None, candidates: Any = None) -> int:
        if kind not in REVIEW_KINDS:
            raise SchemaError("review_queue.kind 只能是 %s：%r" % ("/".join(REVIEW_KINDS), kind))
        cur = self.conn.execute(
            "INSERT INTO review_queue(batch_no,run_id,kind,entity_id,case_id,template_id,field,"
            "label,reason,candidates_json,created_at) VALUES(?,?,?,?,?,?,?,?,?,?,?)",
            (batch_no, run_id, kind, entity_id, case_id, template_id, field, label, reason,
             None if candidates is None else json.dumps(candidates, ensure_ascii=False, default=str),
             now_utc()))
        self.conn.commit()
        return int(cur.lastrowid)

    def open_reviews(self, kind: str | None = None) -> list[dict]:
        sql = "SELECT * FROM review_queue WHERE resolved_at IS NULL"
        args: tuple = ()
        if kind:
            sql += " AND kind=?"
            args = (kind,)
        return [dict(r) for r in self.conn.execute(sql + " ORDER BY id", args)]

    def resolve_review(self, review_id: int, answer: str, *, answered_value: str | None = None,
                       decided_by: str = "user") -> None:
        self.conn.execute("UPDATE review_queue SET answer=?, answered_value=?, decided_by=?,"
                          " resolved_at=? WHERE id=?",
                          (answer, answered_value, decided_by, now_utc(), review_id))
        self.conn.commit()

    # ---- 追溯与汇总 ----------------------------------------------------
    def trace(self, key: str) -> dict:
        facts = [dict(r) for r in self.conn.execute(
            "SELECT f.*, s.name AS source_name, s.path AS source_path FROM fact f"
            " LEFT JOIN source s ON s.id=f.source_id WHERE f.key=? ORDER BY f.id", (key,))]
        for f in facts:
            f["entity"] = self.entity_label(f["entity_id"])
        conflicts = [dict(r) for r in self.conn.execute(
            "SELECT * FROM fact_conflict WHERE key=? ORDER BY id", (key,))]
        uses = [dict(r) for r in self.conn.execute(
            "SELECT * FROM fill_op WHERE field=? ORDER BY id", (key,))]
        reviews = [dict(r) for r in self.conn.execute(
            "SELECT * FROM review_queue WHERE field=? ORDER BY id", (key,))]
        # ⚠️ 不能写成 `'%"key": "%s"' % key`：`%"` 会被当成格式符，直接抛
        # ValueError: unsupported format character。用拼接。
        payload_like = '%"key": "' + str(key) + '"%'
        events = [dict(r) for r in self.conn.execute(
            "SELECT id,batch_no,seq,event_type,occurred_at,actor_id,target,value_text"
            " FROM event WHERE target=? OR payload_json LIKE ? ORDER BY id",
            (key, payload_like))]
        return {"key": key, "facts": facts, "conflicts": conflicts, "fill_ops": uses,
                "reviews": reviews, "events": events}

    def summary(self) -> dict:
        def one(sql: str, args: tuple = ()) -> Any:
            row = self.conn.execute(sql, args).fetchone()
            return row[0] if row else 0

        pending_identity = one("SELECT COUNT(*) FROM entity WHERE identity_incomplete=1")
        return {
            "batch": self.current_batch(),
            "sources": one("SELECT COUNT(*) FROM source"),
            "entities": one("SELECT COUNT(*) FROM entity"),
            "entities_without_key": pending_identity,
            "cases": one('SELECT COUNT(*) FROM "case"'),
            "case_roles": one("SELECT COUNT(*) FROM case_role"),
            "facts_current": one("SELECT COUNT(*) FROM fact WHERE superseded_by IS NULL"),
            "facts_with_value": one("SELECT COUNT(*) FROM fact WHERE superseded_by IS NULL"
                                    " AND value IS NOT NULL AND value<>''"),
            "facts_missing": one("SELECT COUNT(*) FROM fact WHERE superseded_by IS NULL"
                                 " AND (value IS NULL OR value='')"),
            "facts_unassigned": one("SELECT COUNT(*) FROM fact WHERE superseded_by IS NULL"
                                    " AND entity_id IS NULL"),
            "facts_superseded": one("SELECT COUNT(*) FROM fact WHERE superseded_by IS NOT NULL"),
            "conflicts_open": one("SELECT COUNT(*) FROM fact_conflict WHERE resolved=0"),
            "templates": one("SELECT COUNT(*) FROM template"),
            "rules": one("SELECT COUNT(*) FROM template_rule"),
            "rules_required": one("SELECT COUNT(*) FROM template_rule WHERE is_required=1"),
            "fill_ok": one("SELECT COUNT(*) FROM fill_op WHERE kind='field' AND status='ok'"),
            "fill_not_ok": one("SELECT COUNT(*) FROM fill_op WHERE kind='field' AND status<>'ok'"),
            "artifacts": one("SELECT COUNT(*) FROM fill_op WHERE kind='artifact'"),
            "runs_running": one("SELECT COUNT(*) FROM fill_op WHERE kind='run' AND status='running'"),
            "reviews_open": one("SELECT COUNT(*) FROM review_queue WHERE resolved_at IS NULL"),
            "reviews_resolved": one("SELECT COUNT(*) FROM review_queue WHERE resolved_at"
                                    " IS NOT NULL"),
            "events": one("SELECT COUNT(*) FROM event"),
        }

    def current_batch(self) -> str | None:
        for sql in ("SELECT batch_no FROM fact ORDER BY id DESC LIMIT 1",
                    "SELECT batch_no FROM source ORDER BY id DESC LIMIT 1",
                    "SELECT batch_no FROM event ORDER BY id DESC LIMIT 1"):
            r = self.conn.execute(sql).fetchone()
            if r and r[0]:
                return r[0]
        return None


def _norm_rule_for(key: str) -> str:
    """冲突比对用了哪条归一规则（金额→元 / 日期→YYYY-MM-DD / 文本原样，26 §2.8）。"""
    k = str(key)
    if any(w in k for w in ("金额", "额度", "余额", "价款", "利率", "比例")):
        return "amount"
    if any(w in k for w in ("日期", "日", "时间", "期限")):
        return "date"
    return "text"
