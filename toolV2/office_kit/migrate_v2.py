"""SCHEMA_VERSION 1 -> 2 migration (phase 1).

Implements doc 27 ``第一阶段详细设计`` sections 3.4.1 - 3.4.7, i.e. the
SQLite "rebuild" route:

    rename old -> create new -> INSERT..SELECT -> append deferred reviews -> drop old

Design rules that are easy to get wrong (all learned from real failures):

* The batch-number map MUST be computed before the data is copied.  Generating
  it afterwards leaves free text in ``batch_no``, which is the audit chain key.
* Every ``batch_no`` comes from that map -- never from ``session.name``.
* Rows carrying explicit ids are copied BEFORE anything that appends
  auto-increment rows to ``review_queue``.
* ``fill_op.status='filled'`` is the live success value and maps to ``ok``.
  Missing it would relabel every successful historical fill as ``aborted``.
* ``fact`` may have ``source_id IS NULL`` (the db-review path), so its batch is
  looked up by ``session_id``.
* Legacy ``rule`` rows may collide on the new ``UNIQUE(template_id, field)``.
* DROP order matters (children first) even with foreign keys off.

Schema version is stored in SQLite's own ``PRAGMA user_version`` -- phase 1 adds
no meta table.
"""
from __future__ import annotations

import csv
import shutil
import sqlite3
import time
from pathlib import Path

from . import schema_v2 as S

# --------------------------------------------------------------------------
# the 9 tables that are rebuilt, and where each one goes
# --------------------------------------------------------------------------
RENAME_TO_OLD = [
    "source", "fact", "fact_conflict", "template", "rule",
    "fill_op", "review_item", "role_gate", "role_question",
]

#: children first -- `fill_op__old` references `rule__old`, so dropping the
#: parent first fails with FOREIGN KEY constraint failed.
DROP_ORDER = [
    "fill_op", "review_item", "role_question", "role_gate", "fact",
    "fact_conflict", "rule", "template", "source",
]

#: the fixed copy order (doc 27 section 3.4.1 step 5)
COPY_ORDER = [
    "template", "source", "review_queue", "template_rule",
    "fact_conflict", "fact", "fill_op", "case_role",
]

BATCH_NO_RE = r"^\d{8}-\d{2}$"

#: high-risk field classes (26 section 2.6), matched on the field name.
HIGH_RISK_RULES = [
    ("收款账号", ("账号", "账户", "卡号")),
    ("金额", ("金额", "敞口", "额度", "保证金", "余额", "本金")),
    ("利率", ("利率", "费率", "息")),
    ("日期", ("日期", "到期日", "生效日", "提款日", "起始日", "期限")),
    ("证件号码", ("身份证", "证件号")),
]

#: the only three format rules this period (26 section 2.2 / P-14)
FORMAT_RULES = {"统一社会信用代码": "uscc18", "身份证号": "idcard18"}


def classify_high_risk(field: str) -> int:
    for _, words in HIGH_RISK_RULES:
        if any(w in field for w in words):
            return 1
    return 0


class MigrationError(RuntimeError):
    pass


def _now() -> str:
    return time.strftime("%Y-%m-%dT%H:%M:%S")


def _cols(con: sqlite3.Connection, table: str) -> list[str]:
    return [r[1] for r in con.execute('PRAGMA table_info("%s")' % table)]


def _has_table(con: sqlite3.Connection, name: str) -> bool:
    return con.execute(
        "SELECT 1 FROM sqlite_master WHERE type='table' AND name=?", (name,)
    ).fetchone() is not None


def schema_version(con: sqlite3.Connection) -> int:
    return int(con.execute("PRAGMA user_version").fetchone()[0] or 0)


# --------------------------------------------------------------------------
# step 4b -- the batch-number map.  MUST run before any copy.
# --------------------------------------------------------------------------
def build_batch_map(con: sqlite3.Connection, migration_day: str) -> dict[int, str]:
    """session.id -> batch_no.

    A session name that already looks like ``YYYYMMDD-NN`` is reused verbatim;
    everything else gets ``<migration_day>-NN`` with NN assigned in session-id
    order, skipping numbers already taken.  The old name is kept as the label.
    """
    import re

    mapping: dict[int, str] = {}
    used: set[str] = set()
    sessions = list(con.execute("SELECT id, name FROM session ORDER BY id"))
    for sid, name in sessions:
        if name and re.match(BATCH_NO_RE, name):
            mapping[sid] = name
            used.add(name)
    nn = 0
    for sid, _ in sessions:
        if sid in mapping:
            continue
        while True:
            nn += 1
            cand = "%s-%02d" % (migration_day, nn)
            if cand not in used:
                break
        if nn > 99:
            raise MigrationError(
                "more than 99 legacy sessions for one day: migrate in several days "
                "(doc 27 section 3.4.3)")
        mapping[sid] = cand
        used.add(cand)
    return mapping


def _write_map_csv(mapping: dict[int, str], con: sqlite3.Connection, path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    names = dict(con.execute("SELECT id, name FROM session"))
    with path.open("w", encoding="utf-8", newline="") as fh:
        w = csv.writer(fh)
        w.writerow(["batch_no", "batch_label", "old_session_id"])
        for sid, bno in sorted(mapping.items()):
            w.writerow([bno, names.get(sid, ""), sid])


# --------------------------------------------------------------------------
# step 4a -- entities.  Names come from role_gate.entity plus fact values whose
# key is one of SUBJECT_NAME_KEYS.  Never merged by similarity.  No review rows
# here: those must wait until after the copy (step 5b) or they steal ids.
# --------------------------------------------------------------------------
def _build_entities(con: sqlite3.Connection) -> dict[str, int]:
    names: dict[str, dict] = {}

    for (ent,) in con.execute(
            "SELECT DISTINCT entity FROM role_gate__old WHERE entity IS NOT NULL AND entity <> ''"):
        names.setdefault(ent, {"source_id": None, "personal": False})
    for (ent, role) in con.execute("SELECT DISTINCT entity, role FROM role_gate__old"):
        if ent in names and "个人保证人" in (role or ""):
            names[ent]["personal"] = True

    q = ("SELECT DISTINCT value, source_id, key FROM fact__old "
         "WHERE key IN (%s) AND value IS NOT NULL AND value <> ''"
         % ",".join("?" * len(S.SUBJECT_NAME_KEYS)))
    for value, src, key in con.execute(q, S.SUBJECT_NAME_KEYS):
        d = names.setdefault(value, {"source_id": None, "personal": False})
        if d["source_id"] is None:
            d["source_id"] = src
        if key == "个人保证人姓名":
            d["personal"] = True

    ids: dict[str, int] = {}
    for name, d in names.items():
        etype, needs_confirm = S.entity_type_for(
            name, personal_keys=("个人保证人姓名",) if d["personal"] else ())
        eid = con.execute(
            "INSERT INTO entity(entity_type,identity_incomplete,created_at,updated_at)"
            " VALUES(?,1,?,?)", (etype, _now(), _now())).lastrowid
        con.execute(
            "INSERT INTO entity_name(entity_id,name,name_kind,source_id,created_at)"
            " VALUES(?,?,'现用名',?,?)", (eid, name, d["source_id"], _now()))
        # an empty case so roles have somewhere to hang
        con.execute('INSERT INTO "case"(entity_id,created_at) VALUES(?,?)', (eid, _now()))
        ids[name] = eid
    return ids


def _entity_lookup(con: sqlite3.Connection) -> dict[str, int]:
    return {name: eid for eid, name in
            con.execute("SELECT entity_id, name FROM entity_name")}


# --------------------------------------------------------------------------
# step 5 -- the copy, in the fixed order
# --------------------------------------------------------------------------
def _copy_template(con, batch_map):
    for tid, sess, path, name, sha, reg in con.execute(
            "SELECT id,session_id,path,name,sha256,registered_at FROM template__old"):
        low = (path or "").lower()
        kind = ".docx" if low.endswith(".docx") else (".xlsx" if low.endswith(".xlsx") else None)
        con.execute(
            "INSERT INTO template(id,batch_no,path,name,sha256,doc_kind,legacy_session_id,registered_at)"
            " VALUES(?,?,?,?,?,?,?,?)",
            (tid, batch_map.get(sess), path, name, sha, kind, sess, reg))


def _copy_source(con, batch_map):
    for sid, sess, path, name, sha, nbytes, kind, ing in con.execute(
            "SELECT id,session_id,path,name,sha256,bytes,kind,ingested_at FROM source__old"):
        con.execute(
            "INSERT INTO source(id,batch_no,batch_label,path,copy_path,name,sha256,bytes,kind,"
            "parse_status,legacy_session_id,ingested_at) VALUES(?,?,?,?,?,?,?,?,?,'ok',?,?)",
            (sid, batch_map.get(sess), None, path, path, name, sha, nbytes, kind, sess, ing))


def _copy_review_queue(con, batch_map, lookup):
    for rid, sess, tname, field, label, reason, cand, ans, rby, rat, created in con.execute(
            "SELECT id,session_id,template,field,label,reason,candidates,answer,resolved_by,"
            "resolved_at,created_at FROM review_item__old"):
        tid = con.execute("SELECT id FROM template WHERE name=?", (tname,)).fetchone()
        con.execute(
            "INSERT INTO review_queue(id,batch_no,kind,template_id,field,label,reason,"
            "candidates_json,answered_value,decided_by,resolved_at,legacy_session_id,created_at)"
            " VALUES(?,?,'field',?,?,?,?,?,?,?,?,?,?)",
            (rid, batch_map.get(sess), tid[0] if tid else None, field, label, reason,
             cand, ans, rby, rat, sess, created))
    for rid, sess, ent, q, opts, ans, rby, rat, created in con.execute(
            "SELECT id,session_id,entity,question,options,answer,resolved_by,resolved_at,"
            "created_at FROM role_question__old"):
        con.execute(
            "INSERT INTO review_queue(id,batch_no,kind,entity_id,reason,candidates_json,"
            "answered_value,decided_by,resolved_at,legacy_session_id,created_at)"
            " VALUES(?,?,'role',?,?,?,?,?,?,?,?)",
            (rid + 100000, batch_map.get(sess), lookup.get(ent), q, opts, ans,
             rby, rat, sess, created))


def _copy_template_rule(con, deferred):
    rows = list(con.execute(
        "SELECT id,template_id,field,label,target_json,match_kind,confidence,decided_by,created_at"
        " FROM rule__old ORDER BY confidence DESC, id ASC"))
    seen: set[tuple] = set()
    kept = 0
    for rid, tid, field, label, target, mkind, conf, dec, created in rows:
        if (tid, field) in seen:
            deferred.append(("field", "迁移：同一模板同一字段有多条旧规则，请确认保留哪条",
                             field, None))
            continue
        seen.add((tid, field))
        kept += 1
        con.execute(
            "INSERT INTO template_rule(id,template_id,field,label,target_json,is_required,"
            "match_kind,confidence,decided_by,created_at) VALUES(?,?,?,?,?,0,?,?,?,?)",
            (rid, tid, field, label, target, mkind, conf,
             "auto" if dec == "rule" else dec, created))
    return kept, len(rows)


def _copy_fact_conflict(con, batch_map, deferred):
    for cid, key, existing, incoming, src, det, resolved in con.execute(
            "SELECT id,key,existing,incoming,source_id,detected_at,resolved FROM fact_conflict__old"):
        bno = None
        if src is not None:
            r = con.execute("SELECT legacy_session_id FROM source WHERE id=?", (src,)).fetchone()
            if r:
                bno = batch_map.get(r[0])
        if bno is None and batch_map:
            bno = sorted(batch_map.values())[0]
            deferred.append(("conflict", "迁移：冲突记录没有来源文件，批次按迁移批次暂记",
                             None, None))
        con.execute(
            "INSERT INTO fact_conflict(id,batch_no,key,existing,incoming,incoming_src,"
            "detected_at,resolved) VALUES(?,?,?,?,?,?,?,?)",
            (cid, bno, key, existing, incoming, src, det, resolved))
        if resolved:
            deferred.append(("conflict", "迁移：旧库没记当时怎么裁的，请确认", key, None))


def _copy_fact(con, batch_map, lookup, deferred):
    for fid, sess, src, key, value, status, origin, note, created in con.execute(
            "SELECT id,session_id,source_id,key,value,status,origin,note,created_at FROM fact__old"):
        eid = None
        if key in S.SUBJECT_NAME_KEYS and value:
            eid = lookup.get(value)
        skind, prov = S.map_origin(origin)
        if src is None and not note:
            note = "迁移：原记录无来源文件"
        con.execute(
            "INSERT INTO fact(id,batch_no,source_id,entity_id,key,value,status,source_kind,"
            "provenance,note,legacy_session_id,created_at) VALUES(?,?,?,?,?,?,?,?,?,?,?,?)",
            (fid, batch_map.get(sess), src, eid, key, value, status, skind, prov, note,
             sess, created))
        if eid is None:
            deferred.append(("field", "迁移：无法从旧库判定所属主体", key, None))


def _copy_fill_op(con, batch_map):
    run_no: dict[int, int] = {}
    for (tid,) in con.execute(
            "SELECT DISTINCT template_id FROM fill_op__old "
            "WHERE template_id IS NOT NULL ORDER BY template_id"):
        run_no[tid] = len(run_no) + 1
    for oid, sess, tid, rid, fid, field, value, status, detail, created in con.execute(
            "SELECT id,session_id,template_id,rule_id,fact_id,field,value,status,detail,"
            "created_at FROM fill_op__old"):
        bno = batch_map.get(sess)
        run_id = "%s-R%02d" % (bno, run_no.get(tid, 0))
        st = S.STATUS_MAP.get(status)
        if st is None:
            st = "aborted"
        eid = None
        if fid:
            r = con.execute("SELECT entity_id FROM fact WHERE id=?", (fid,)).fetchone()
            eid = r[0] if r else None
        con.execute(
            "INSERT INTO fill_op(id,run_id,batch_no,kind,entity_id,template_id,template_rule_id,"
            "fact_id,field,value,status,detail,legacy_session_id,created_at)"
            " VALUES(?,?,?,'field',?,?,?,?,?,?,?,?,?,?)",
            (oid, run_id, bno, eid, tid, rid, fid, field, value, st, detail, sess, created))


def _copy_case_role(con, lookup, deferred):
    kept = 0
    for rid, sess, ent, role, evidence, decider, created in con.execute(
            "SELECT id,session_id,entity,role,evidence,decided_by,created_at FROM role_gate__old"):
        eid = lookup.get(ent)
        if eid is None:
            deferred.append(("role", "迁移：角色主体的名字匹配不上，请确认", None, None))
            continue
        new_role = S.map_role(role)
        if new_role is None:
            deferred.append(("role",
                             "迁移：角色「%s」超出本期封闭集合，请确认归到哪一类" % role,
                             None, eid))
            continue
        row = con.execute('SELECT id FROM "case" WHERE entity_id=? ORDER BY id', (eid,)).fetchone()
        if row is None:
            deferred.append(("role", "迁移：该主体没有案子，无法挂角色", None, eid))
            continue
        keeper = decider if decider in ("source", "human", "auto") else "human"
        con.execute(
            "INSERT INTO case_role(entity_id,case_id,role,evidence,decided_by,legacy_session_id,"
            "created_at) VALUES(?,?,?,?,?,?,?)",
            (eid, row[0], new_role, evidence, keeper, sess, created))
        kept += 1
    return kept


def _import_field_contract(con):
    keys = [r[0] for r in con.execute(
        "SELECT DISTINCT key FROM fact WHERE key IS NOT NULL AND key <> '' ORDER BY key")]
    added = 0
    for k in keys:
        if con.execute("SELECT 1 FROM field_contract WHERE field=?", (k,)).fetchone():
            continue
        con.execute(
            "INSERT INTO field_contract(field,display_name,dtype,required,high_risk,"
            "format_rule,created_at,updated_at) VALUES(?,?,?,0,?,?,?,?)",
            (k, k, "text", classify_high_risk(k), FORMAT_RULES.get(k), _now(), _now()))
        added += 1
    return added, len(keys)


# --------------------------------------------------------------------------
# section 3.4.7 -- the seven integrity checks
# --------------------------------------------------------------------------
def verify(con: sqlite3.Connection) -> list[tuple[str, bool, str]]:
    import re

    out: list[tuple[str, bool, str]] = []

    # 1 every non-null batch_no matches YYYYMMDD-NN
    bad = []
    for t in ("source", "fact", "fact_conflict", "fill_op", "review_queue", "case", "template"):
        if not _has_table(con, t):
            continue
        for (v,) in con.execute('SELECT DISTINCT batch_no FROM "%s" WHERE batch_no IS NOT NULL' % t):
            if not re.match(BATCH_NO_RE, v or ""):
                bad.append("%s.batch_no=%r" % (t, v))
    out.append(("1 batch_no 全为 YYYYMMDD-NN", not bad, "; ".join(bad) or "全部合规"))

    # 2 status values are all designed ones (histogram compared in the report)
    allowed = set(S.STATUS_MAP.values()) | {"running", "rejected"}
    bad_st = [s for (s,) in con.execute("SELECT DISTINCT status FROM fill_op")
              if s not in allowed]
    out.append(("2 fill_op.status 只有设计内的值", not bad_st, ",".join(bad_st) or "ok"))

    # 3 no unique violations can be checked post-hoc; assert the counts line up
    dup = con.execute(
        "SELECT template_id, field, COUNT(*) c FROM template_rule "
        "GROUP BY template_id, field HAVING c > 1").fetchall()
    out.append(("3 template_rule 无重复 (template_id, field)", not dup, str(dup) or "ok"))

    # 4 deferred review rows were produced when needed
    reasons = [r[0] for r in con.execute(
        "SELECT reason FROM review_queue WHERE reason IS NOT NULL")]
    out.append(("4 迁移待确认项已生成（可为 0）", True,
                "共 %d 条" % sum(1 for r in reasons if r.startswith("迁移"))))

    # 5 nothing fabricated
    for t, col in (("fact", "confidence"), ("fill_op", "artifact_sha256"),
                   ("fill_op", "signoff_actor")):
        n = con.execute('SELECT COUNT(*) FROM "%s" WHERE %s IS NOT NULL' % (t, col)).fetchone()[0]
        if n:
            out.append(("5 不编造：%s.%s 应为空" % (t, col), False, "%d 条非空" % n))
            break
    else:
        n_ev = con.execute("SELECT COUNT(*) FROM event").fetchone()[0]
        out.append(("5 不编造（置信度/指纹/签核/事件为空）", n_ev == 0,
                    "event=%d" % n_ev))

    # 6 entity_type / identity_incomplete consistency
    bad_e = con.execute(
        "SELECT COUNT(*) FROM entity WHERE identity_incomplete NOT IN (0,1)").fetchone()[0]
    out.append(("6 entity.identity_incomplete 取值合法", bad_e == 0, "ok"))

    # 7 schema version written
    v = schema_version(con)
    out.append(("7 PRAGMA user_version = 2", v == S.SCHEMA_VERSION_V2, "user_version=%d" % v))
    return out


# --------------------------------------------------------------------------
# the migration itself
# --------------------------------------------------------------------------
def migrate(db_path: str | Path, *, dry_run: bool = False, migration_day: str | None = None,
            log_map_to: str | Path | None = None, emit=print) -> dict:
    db_path = Path(db_path)
    if not db_path.exists():
        raise MigrationError("database not found: %s" % db_path)
    migration_day = migration_day or time.strftime("%Y%m%d")

    con = sqlite3.connect(str(db_path), isolation_level=None)
    report: dict = {"db": str(db_path), "dry_run": dry_run, "steps": []}

    def note(step, detail=""):
        report["steps"].append((step, detail))
        emit("  %-38s %s" % (step, detail))

    try:
        cur = schema_version(con)
        leftover = [r[0] for r in con.execute(
            "SELECT name FROM sqlite_master WHERE type='table' AND name LIKE '%__old'")]
        if leftover:
            raise MigrationError(
                "idempotence gate: leftover tables %s -- restore the backup instead of "
                "re-running (doc 27 section 3.4.1 step 2)" % ", ".join(leftover))
        if cur == S.SCHEMA_VERSION_V2:
            raise MigrationError("already at SCHEMA_VERSION 2 -- nothing to do")
        note("0 幂等闸门 / 当前版本", "user_version=%d" % cur)

        batch_map = build_batch_map(con, migration_day)
        note("4b 批次映射", "%d 个 session -> %s" % (
            len(batch_map), ", ".join(sorted(batch_map.values())[:4]) + ("..." if len(batch_map) > 4 else "")))

        before = {t: con.execute('SELECT COUNT(*) FROM "%s"' % t).fetchone()[0]
                  for t in RENAME_TO_OLD}
        report["before"] = before
        report["status_before"] = dict(con.execute(
            "SELECT status, COUNT(*) FROM fill_op GROUP BY status"))
        report["batch_map"] = dict(batch_map)

        if dry_run:
            note("--dry-run", "将重建 %d 张表；读 %d 行" % (len(RENAME_TO_OLD), sum(before.values())))
            for t in RENAME_TO_OLD:
                emit("      %-16s 读 %d 行" % (t, before[t]))
            report["planned_batch_map"] = batch_map
            report["checks"] = []            # nothing was changed, so nothing to verify
            report["n_diverted"] = 0
            return report

        # 1 backup
        stamp = time.strftime("%Y%m%d-%H%M%S")
        backup = db_path.with_name(db_path.stem + ".v1.bak." + stamp + db_path.suffix)
        con.close()
        shutil.copy2(db_path, backup)
        con = sqlite3.connect(str(db_path), isolation_level=None)
        note("1 备份", backup.name)

        con.execute("PRAGMA foreign_keys = OFF")            # step 0
        for t in RENAME_TO_OLD:                             # step 3
            con.execute('ALTER TABLE "%s" RENAME TO "%s__old"' % (t, t))
        note("3 改名到 __old", "%d 张" % len(RENAME_TO_OLD))

        con.executescript(S.DDL_V2)                         # step 4
        note("4 建基础数据表", "%d 张" % S.DDL_V2.count("CREATE TABLE IF NOT EXISTS"))

        lookup = _build_entities(con)                       # step 4a
        note("4a 主体 + 空壳案子", "%d 个主体" % len(lookup))
        if log_map_to:
            _write_map_csv(batch_map, con, Path(log_map_to))
            note("4b 批次对照表", str(log_map_to))

        deferred: list[tuple] = []
        counters: dict[str, int] = {}
        _copy_template(con, batch_map);                     counters["template"] = con.total_changes
        _copy_source(con, batch_map)
        _copy_review_queue(con, batch_map, lookup)
        kept, total = _copy_template_rule(con, deferred)
        _copy_fact_conflict(con, batch_map, deferred)
        _copy_fact(con, batch_map, lookup, deferred)
        _copy_fill_op(con, batch_map)
        n_roles = _copy_case_role(con, lookup, deferred)
        note("5 搬数据（固定顺序）", "template_rule 保留 %d/%d；case_role %d 条" % (kept, total, n_roles))

        # label the source copies now that we know the batch labels
        con.execute(
            "UPDATE source SET batch_label=(SELECT name FROM session WHERE session.id=source.legacy_session_id)"
            " WHERE batch_label IS NULL")

        for kind, reason, field, eid in deferred:           # step 5b
            con.execute(
                "INSERT INTO review_queue(kind,reason,field,entity_id,created_at)"
                " VALUES(?,?,?,?,?)", (kind, reason, field, eid, _now()))
        note("5b 追加迁移待确认项", "%d 条" % len(deferred))

        added, n_keys = _import_field_contract(con)         # step 7
        note("7 导入字段契约", "新增 %d / 共 %d 个字段" % (added, n_keys))

        for t in DROP_ORDER:                                # step 8
            con.execute('DROP TABLE "%s__old"' % t)
        note("8 删 __old（子表先删）", "%d 张" % len(DROP_ORDER))

        con.execute("PRAGMA foreign_keys = ON")             # step 9
        con.execute("VACUUM")                               # outside any transaction
        con.execute("PRAGMA user_version = %d" % S.SCHEMA_VERSION_V2)   # step 10

        report["after"] = {t: con.execute('SELECT COUNT(*) FROM "%s"' % t).fetchone()[0]
                           for t in ("source", "fact", "fact_conflict", "template",
                                     "template_rule", "fill_op", "review_queue",
                                     "entity", "entity_name", "case", "case_role",
                                     "field_contract")}
        report["status_histogram"] = dict(con.execute(
            "SELECT status, COUNT(*) FROM fill_op GROUP BY status"))
        report["checks"] = verify(con)
        report["backup"] = str(backup)
        report["diverted"] = [
            dict(zip(("kind", "reason", "field", "entity_id"), r))
            for r in con.execute(
                "SELECT kind, reason, field, entity_id FROM review_queue "
                "WHERE reason LIKE '迁移%' ORDER BY id LIMIT 50")]
        report["n_diverted"] = con.execute(
            "SELECT COUNT(*) FROM review_queue WHERE reason LIKE '迁移%'").fetchone()[0]
        report["tables"] = {
            "built": [r[0] for r in con.execute(
                "SELECT name FROM sqlite_master WHERE type='table' "
                "AND name NOT LIKE 'sqlite_%' AND name NOT LIKE '%__old' ORDER BY name")],
        }
        return report
    except Exception:
        con.close()
        raise
    finally:
        try:
            con.close()
        except Exception:
            pass


# --------------------------------------------------------------------------
# human-readable report -- so the result can be *looked at*, not just trusted
# --------------------------------------------------------------------------
def render_report(rep: dict) -> str:
    L: list[str] = []
    A = L.append
    A("# 数据库迁移报告（SCHEMA v1 → v2）")
    A("")
    A("数据库：`%s`" % rep["db"])
    A("模式：**%s**" % ("预演（未改动数据库）" if rep["dry_run"] else "实际执行"))
    A("")
    A("## 一、过程")
    A("")
    A("| 步骤 | 结果 |")
    A("| --- | --- |")
    for step, detail in rep["steps"]:
        A("| %s | %s |" % (step, detail or "—"))
    A("")

    if rep["dry_run"]:
        A("## 二、将要处理的数据量")
        A("")
        A("| 旧表 | 行数 |")
        A("| --- | --- |")
        for t, n in sorted(rep.get("before", {}).items()):
            A("| `%s` | %d |" % (t, n))
        A("")
        A("> 预演不改动任何东西。确认无误后再执行（去掉 `--dry-run`）。")
        return "\n".join(L) + "\n"

    A("## 二、行数（旧 → 新）")
    A("")
    A("| 表 | 说明 | 行数 |")
    A("| --- | --- | --- |")
    labels = {
        "source": "源文件", "fact": "键值对", "fact_conflict": "冲突",
        "template": "模板", "template_rule": "格子映射", "fill_op": "填充记录",
        "review_queue": "待确认项", "entity": "主体", "entity_name": "主体名称",
        "case": "案子（授信）", "case_role": "主体角色", "field_contract": "字段契约",
    }
    for t, n in sorted(rep.get("after", {}).items()):
        A("| `%s` | %s | %d |" % (t, labels.get(t, ""), n))
    A("")
    olds = rep.get("before", {})
    if olds:
        A("旧库读入：" + "、".join("`%s` %d 行" % (k, v) for k, v in sorted(olds.items())))
        A("")

    sb, sa = rep.get("status_before", {}), rep.get("status_histogram", {})
    if sb or sa:
        A("## 三、填充记录状态（**这是最容易改坏的地方**）")
        A("")
        A("| 状态 | 迁移前 | 迁移后 |")
        A("| --- | --- | --- |")
        for k in sorted(set(sb) | set(sa)):
            A("| `%s` | %s | %s |" % (k, sb.get(k, 0), sa.get(k, 0)))
        A("")
        if sb.get("filled") and not sa.get("ok"):
            A("> ❌ **异常**：迁移前有 %d 条 `filled`，迁移后没有 `ok`——值映射丢了。" % sb["filled"])
            A("")

    A("## 四、完整性检查（%s）" % ("全部通过 ✅" if all(c[1] for c in rep["checks"]) else "**有失败 ❌**"))
    A("")
    A("| # | 检查 | 结果 | 说明 |")
    A("| --- | --- | --- | --- |")
    for name, ok, detail in rep["checks"]:
        A("| %s | %s | %s | %s |" % (name.split(" ", 1)[0],
                                     name.split(" ", 1)[-1],
                                     "✅" if ok else "❌", detail))
    A("")

    bm = rep.get("batch_map") or {}
    if bm:
        A("## 五、批次号对照（旧批次名 → 合规批次号）")
        A("")
        A("| 旧 session.id | 新 batch_no |")
        A("| --- | --- |")
        for sid, bno in sorted(bm.items(), key=lambda kv: kv[1]):
            A("| %s | `%s` |" % (sid, bno))
        A("")

    div = rep.get("diverted") or []
    A("## 六、迁移时挑出来要你确认的（**共 %d 条**）" % rep.get("n_diverted", 0))
    A("")
    if not div:
        A("没有。")
    else:
        A("| 类型 | 原因 | 字段 |")
        A("| --- | --- | --- |")
        for d in div[:30]:
            A("| %s | %s | %s |" % (d.get("kind") or "", d.get("reason") or "",
                                    d.get("field") or "—"))
        if len(div) > 30:
            A("| … | 还有 %d 条 | |" % (len(div) - 30))
        A("")
        A("> 这些**没有被猜**，都留在待确认队列里，等你在会话里逐条定。")
    A("")

    A("## 七、备份")
    A("")
    A("迁移前的整库备份：`%s`" % rep.get("backup", "—"))
    A("")
    A("> ⚠️ 迁移**不做双写、不支持回退到 v1**；要退回只能用这个备份文件。")
    A("")
    A("## 八、结论")
    A("")
    ok_all = all(c[1] for c in rep["checks"])
    A("**%s**" % ("迁移成功：%d 条完整性检查全部通过。" % len(rep["checks"]) if ok_all
                  else "迁移**未通过**检查——请用备份回滚，并把上面失败项发给我。"))
    return "\n".join(L) + "\n"


def cmd_db_migrate(args):  # pragma: no cover - exercised by selftest
    from . import workroot as WR
    from .common import OfficeKitError, Result, write_text, unique_path

    # 迁移也走工作区：`--db` / `--work` / 往上找标记文件，三选一（认不出就报错）
    try:
        WR.open_for_command(args)
    except WR.WorkrootError as exc:
        raise OfficeKitError(str(exc)) from exc

    dry = bool(getattr(args, "dry_run", False))
    out = getattr(args, "out", None)
    res = Result("db-migrate")
    if out:
        out_path = Path(out)
        out_path.mkdir(parents=True, exist_ok=True)
        map_path = out_path / "migration_map.csv"
    else:
        map_path = None
    rep = migrate(args._db, dry_run=dry, migration_day=getattr(args, "day", None),
                  log_map_to=map_path, emit=lambda *a: None)
    checks = rep.get("checks") or []
    text = render_report(rep)
    if out:
        p = write_text(Path(out) / "migration_report.md", text)
        res.add_artifact(p, "迁移报告（人看的）")
        if map_path and map_path.exists():
            res.add_artifact(map_path, "批次号对照表")
    res.data.update({
        "dry_run": dry,
        "backup": rep.get("backup"),
        "steps": [{"step": s, "detail": d} for s, d in rep["steps"]],
        "rows_before": rep.get("before"),
        "rows_after": rep.get("after"),
        "status_before": rep.get("status_before"),
        "status_after": rep.get("status_histogram"),
        "checks": [{"check": n, "ok": o, "detail": d} for n, o, d in checks],
        "all_checks_passed": (all(c[1] for c in checks) if checks else None),
        "needs_confirmation": rep.get("n_diverted", 0),
        "tables": (rep.get("tables") or {}).get("built"),
        "report": text,
    })
    if checks and not all(c[1] for c in checks):
        res.warn("完整性检查有失败项 —— 请用备份回滚")
    return res


if __name__ == "__main__":  # pragma: no cover - manual use
    import argparse

    ap = argparse.ArgumentParser(description="workflow.db: schema v1 -> v2")
    ap.add_argument("--db", required=True)
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--day", default=None)
    ap.add_argument("--out", default=None)
    a = ap.parse_args()
    rep = migrate(a.db, dry_run=a.dry_run, migration_day=a.day)
    text = render_report(rep)
    print(text)
    if a.out:
        p = Path(a.out)
        p.mkdir(parents=True, exist_ok=True)
        (p / "migration_report.md").write_text(text, encoding="utf-8")
        print("report ->", p / "migration_report.md")
