"""Persistent, traceable workflow store for form filling.

Why a database instead of loose JSON files
------------------------------------------
The user's requirement is that information stays *traceable* and *reusable across
sessions*: where a value came from, which rule filled which document, what was
skipped and why, and what a human decided. A single SQLite file gives all of that
one queryable home, survives across conversations, and needs no server.

Tables
------
session         one row per run (who/when/which command), the audit anchor
source          an ingested file, with hash so re-ingest is detected
fact            one key/value extracted from a source, with provenance
fact_conflict   when a key was seen with a different value
template        a target form plus its bytes-hash
rule            label/cell -> fact mapping for a template
fill_op         one attempt to write one fact into one template
review_item     anything the program refused to decide, with its resolution
"""
from __future__ import annotations

import hashlib
import json
import sqlite3
from datetime import datetime
from pathlib import Path
from typing import Any


SCHEMA_VERSION = 1

DDL = """
CREATE TABLE IF NOT EXISTS session (
    id           INTEGER PRIMARY KEY AUTOINCREMENT,
    name         TEXT NOT NULL,
    created_at   TEXT NOT NULL,
    note         TEXT
);

CREATE TABLE IF NOT EXISTS source (
    id           INTEGER PRIMARY KEY AUTOINCREMENT,
    session_id   INTEGER REFERENCES session(id),
    path         TEXT NOT NULL,
    name         TEXT NOT NULL,
    sha256       TEXT NOT NULL,
    bytes        INTEGER NOT NULL,
    kind         TEXT,
    ingested_at  TEXT NOT NULL,
    UNIQUE(path, sha256)
);

CREATE TABLE IF NOT EXISTS fact (
    id           INTEGER PRIMARY KEY AUTOINCREMENT,
    session_id   INTEGER REFERENCES session(id),
    source_id    INTEGER REFERENCES source(id),
    key          TEXT NOT NULL,
    value        TEXT,
    status       TEXT NOT NULL DEFAULT 'ok',
    origin       TEXT,
    note         TEXT,
    created_at   TEXT NOT NULL
);
CREATE INDEX IF NOT EXISTS idx_fact_key ON fact(key);

CREATE TABLE IF NOT EXISTS fact_conflict (
    id           INTEGER PRIMARY KEY AUTOINCREMENT,
    key          TEXT NOT NULL,
    existing     TEXT,
    incoming     TEXT,
    source_id    INTEGER,
    detected_at  TEXT NOT NULL,
    resolved     INTEGER NOT NULL DEFAULT 0
);

CREATE TABLE IF NOT EXISTS template (
    id           INTEGER PRIMARY KEY AUTOINCREMENT,
    session_id   INTEGER REFERENCES session(id),
    path         TEXT NOT NULL,
    name         TEXT NOT NULL,
    sha256       TEXT NOT NULL,
    registered_at TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS rule (
    id           INTEGER PRIMARY KEY AUTOINCREMENT,
    template_id  INTEGER REFERENCES template(id),
    field        TEXT NOT NULL,
    label        TEXT,
    target_json  TEXT NOT NULL,
    match_kind   TEXT NOT NULL DEFAULT 'exact',
    confidence   REAL,
    decided_by   TEXT,
    created_at   TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS fill_op (
    id           INTEGER PRIMARY KEY AUTOINCREMENT,
    session_id   INTEGER REFERENCES session(id),
    template_id  INTEGER REFERENCES template(id),
    rule_id      INTEGER REFERENCES rule(id),
    fact_id      INTEGER REFERENCES fact(id),
    field        TEXT NOT NULL,
    value        TEXT,
    status       TEXT NOT NULL,
    detail       TEXT,
    created_at   TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS review_item (
    id           INTEGER PRIMARY KEY AUTOINCREMENT,
    session_id   INTEGER REFERENCES session(id),
    template     TEXT,
    field        TEXT,
    label        TEXT,
    reason       TEXT NOT NULL,
    candidates   TEXT,
    answer       TEXT,
    resolved_by  TEXT,
    resolved_at  TEXT,
    created_at   TEXT NOT NULL
);
CREATE INDEX IF NOT EXISTS idx_review_open ON review_item(resolved_at);

-- A role gate: entity -> the role it plays (借款人/保证人/法定代表人...).
-- Filling a form asserts WHO someone is on that form, so an unconfirmed role is a
-- business decision, not a formatting detail. Rows are written either because the
-- source document states the role explicitly (auto) or because a human answered
-- the question (human).
CREATE TABLE IF NOT EXISTS role_gate (
    id           INTEGER PRIMARY KEY AUTOINCREMENT,
    session_id   INTEGER REFERENCES session(id),
    entity       TEXT NOT NULL,
    role         TEXT NOT NULL,
    evidence     TEXT,
    decided_by   TEXT NOT NULL DEFAULT 'human',
    created_at   TEXT NOT NULL,
    UNIQUE(entity, role)
);

-- Open questions about roles. Kept separate from ordinary field reviews because an
-- unanswered role blocks the whole run rather than just one field.
CREATE TABLE IF NOT EXISTS role_question (
    id           INTEGER PRIMARY KEY AUTOINCREMENT,
    session_id   INTEGER REFERENCES session(id),
    entity       TEXT NOT NULL,
    question     TEXT NOT NULL,
    options      TEXT,
    answer       TEXT,
    resolved_by  TEXT,
    resolved_at  TEXT,
    created_at   TEXT NOT NULL
);
"""


def _now() -> str:
    return datetime.now().isoformat(timespec="seconds")


def sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as fh:
        for block in iter(lambda: fh.read(1 << 20), b""):
            h.update(block)
    return h.hexdigest()


class Store:
    """Thin, explicit wrapper over the SQLite workflow database."""

    def __init__(self, path: str | Path):
        self.path = Path(path)
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self.conn = sqlite3.connect(str(self.path))
        self.conn.row_factory = sqlite3.Row
        self.conn.executescript(DDL)
        self.conn.execute(f"PRAGMA user_version = {SCHEMA_VERSION}")
        self.conn.commit()

    def close(self) -> None:
        self.conn.close()

    def __enter__(self) -> "Store":
        return self

    def __exit__(self, *exc) -> None:
        self.close()

    # ---- session -------------------------------------------------------
    def start_session(self, name: str, note: str | None = None) -> int:
        cur = self.conn.execute(
            "INSERT INTO session (name, created_at, note) VALUES (?,?,?)", (name, _now(), note)
        )
        self.conn.commit()
        return int(cur.lastrowid)

    # ---- sources and facts --------------------------------------------
    def ingest_source(self, path: Path, session_id: int, kind: str | None = None) -> tuple[int, bool]:
        """Record a source file. Returns (source_id, already_ingested)."""
        digest = sha256_file(path)
        row = self.conn.execute(
            "SELECT id FROM source WHERE path=? AND sha256=?", (str(path), digest)
        ).fetchone()
        if row:
            return int(row["id"]), True
        cur = self.conn.execute(
            "INSERT INTO source (session_id, path, name, sha256, bytes, kind, ingested_at)"
            " VALUES (?,?,?,?,?,?,?)",
            (session_id, str(path), path.name, digest, path.stat().st_size, kind, _now()),
        )
        self.conn.commit()
        return int(cur.lastrowid), False

    def put_fact(self, session_id: int, source_id: int | None, key: str, value: Any,
                 *, status: str = "ok", origin: str | None = None, note: str | None = None,
                 overwrite: bool = False) -> int:
        """Store a fact, recording a conflict instead of silently overwriting."""
        existing = self.conn.execute(
            "SELECT id, value FROM fact WHERE key=? ORDER BY id DESC LIMIT 1", (key,)
        ).fetchone()
        text = None if value is None else str(value)
        if existing is not None and not overwrite:
            if (existing["value"] or "") != (text or ""):
                self.conn.execute(
                    "INSERT INTO fact_conflict (key, existing, incoming, source_id, detected_at)"
                    " VALUES (?,?,?,?,?)",
                    (key, existing["value"], text, source_id, _now()),
                )
                self.conn.commit()
                return int(existing["id"])
        cur = self.conn.execute(
            "INSERT INTO fact (session_id, source_id, key, value, status, origin, note, created_at)"
            " VALUES (?,?,?,?,?,?,?,?)",
            (session_id, source_id, key, text, status, origin, note, _now()),
        )
        self.conn.commit()
        return int(cur.lastrowid)

    def facts(self, *, include_missing: bool = True) -> dict[str, dict[str, Any]]:
        """Current value per key (latest row wins)."""
        rows = self.conn.execute("SELECT * FROM fact ORDER BY id").fetchall()
        out: dict[str, dict[str, Any]] = {}
        for r in rows:
            out[r["key"]] = dict(r)
        if not include_missing:
            out = {k: v for k, v in out.items() if v.get("value")}
        return out

    def open_conflicts(self) -> list[dict[str, Any]]:
        return [dict(r) for r in self.conn.execute(
            "SELECT * FROM fact_conflict WHERE resolved=0 ORDER BY id").fetchall()]

    # ---- templates and rules ------------------------------------------
    def register_template(self, path: Path, session_id: int) -> int:
        digest = sha256_file(path)
        row = self.conn.execute("SELECT id FROM template WHERE path=?", (str(path),)).fetchone()
        if row:
            self.conn.execute(
                "UPDATE template SET sha256=?, registered_at=? WHERE id=?",
                (digest, _now(), row["id"]),
            )
            self.conn.commit()
            return int(row["id"])
        cur = self.conn.execute(
            "INSERT INTO template (session_id, path, name, sha256, registered_at) VALUES (?,?,?,?,?)",
            (session_id, str(path), path.name, digest, _now()),
        )
        self.conn.commit()
        return int(cur.lastrowid)

    def add_rule(self, template_id: int, field: str, label: str | None, target: dict[str, Any],
                 *, match_kind: str = "exact", confidence: float | None = None,
                 decided_by: str = "human") -> int:
        cur = self.conn.execute(
            "INSERT INTO rule (template_id, field, label, target_json, match_kind, confidence,"
            " decided_by, created_at) VALUES (?,?,?,?,?,?,?,?)",
            (template_id, field, label, json.dumps(target, ensure_ascii=False),
             match_kind, confidence, decided_by, _now()),
        )
        self.conn.commit()
        return int(cur.lastrowid)

    def rules_for(self, template_path: str) -> list[dict[str, Any]]:
        rows = self.conn.execute(
            "SELECT r.*, t.path AS template_path, t.name AS template_name FROM rule r"
            " JOIN template t ON t.id = r.template_id WHERE t.path=? ORDER BY r.id",
            (str(template_path),),
        ).fetchall()
        out = []
        for r in rows:
            d = dict(r)
            d["target"] = json.loads(d.pop("target_json"))
            out.append(d)
        return out

    def mapping_for(self, template_path: str) -> dict[str, Any]:
        rules = self.rules_for(template_path)
        return {"fields": [{"field": r["field"], "label": r["label"], "target": r["target"]}
                           for r in rules]}

    # ---- run records ---------------------------------------------------
    def record_fill(self, session_id: int, template_id: int | None, rule_id: int | None,
                    fact_id: int | None, field: str, value: str | None,
                    status: str, detail: str = "") -> None:
        self.conn.execute(
            "INSERT INTO fill_op (session_id, template_id, rule_id, fact_id, field, value,"
            " status, detail, created_at) VALUES (?,?,?,?,?,?,?,?,?)",
            (session_id, template_id, rule_id, fact_id, field, value, status, detail, _now()),
        )
        self.conn.commit()

    def add_review(self, session_id: int, template: str | None, field: str | None,
                   label: str | None, reason: str, candidates: Any = None) -> int:
        cur = self.conn.execute(
            "INSERT INTO review_item (session_id, template, field, label, reason, candidates,"
            " created_at) VALUES (?,?,?,?,?,?,?)",
            (session_id, template, field, label, reason,
             json.dumps(candidates, ensure_ascii=False) if candidates is not None else None,
             _now()),
        )
        self.conn.commit()
        return int(cur.lastrowid)

    def open_reviews(self) -> list[dict[str, Any]]:
        return [dict(r) for r in self.conn.execute(
            "SELECT * FROM review_item WHERE resolved_at IS NULL ORDER BY id").fetchall()]

    def resolve_review(self, review_id: int, answer: str, resolved_by: str = "human") -> None:
        self.conn.execute(
            "UPDATE review_item SET answer=?, resolved_by=?, resolved_at=? WHERE id=?",
            (answer, resolved_by, _now(), review_id),
        )
        self.conn.commit()

    # ---- role gate -----------------------------------------------------
    def set_role(self, session_id: int, entity: str, role: str, *,
                 evidence: str | None = None, decided_by: str = "human") -> None:
        self.conn.execute(
            "INSERT OR REPLACE INTO role_gate (session_id, entity, role, evidence,"
            " decided_by, created_at) VALUES (?,?,?,?,?,?)",
            (session_id, entity, role, evidence, decided_by, _now()),
        )
        self.conn.commit()

    def roles(self) -> list[dict[str, Any]]:
        return [dict(r) for r in self.conn.execute(
            "SELECT * FROM role_gate ORDER BY entity, role").fetchall()]

    def role_of(self, entity: str) -> list[str]:
        return [r["role"] for r in self.conn.execute(
            "SELECT role FROM role_gate WHERE entity=? ORDER BY role", (entity,)).fetchall()]

    def ask_role(self, session_id: int, entity: str, question: str,
                 options: Any = None) -> int:
        cur = self.conn.execute(
            "INSERT INTO role_question (session_id, entity, question, options, created_at)"
            " VALUES (?,?,?,?,?)",
            (session_id, entity, question,
             json.dumps(options, ensure_ascii=False) if options is not None else None, _now()),
        )
        self.conn.commit()
        return int(cur.lastrowid)

    def open_role_questions(self) -> list[dict[str, Any]]:
        return [dict(r) for r in self.conn.execute(
            "SELECT * FROM role_question WHERE resolved_at IS NULL ORDER BY id").fetchall()]

    def resolve_role_question(self, qid: int, answer: str, *, entity: str | None = None,
                              resolved_by: str = "human", session_id: int | None = None) -> None:
        row = self.conn.execute(
            "SELECT entity FROM role_question WHERE id=?", (qid,)).fetchone()
        self.conn.execute(
            "UPDATE role_question SET answer=?, resolved_by=?, resolved_at=? WHERE id=?",
            (answer, resolved_by, _now(), qid),
        )
        self.conn.commit()
        target = entity or (row["entity"] if row else None)
        if target and answer:
            self.set_role(session_id or 0, target, answer,
                          evidence=f"来自 role_question #{qid}", decided_by=resolved_by)

    def role_gate_ok(self) -> bool:
        """True when nothing is left to confirm about who is who."""
        return not self.open_role_questions()

    # ---- reporting -----------------------------------------------------
    def trace(self, key: str) -> dict[str, Any]:
        """Everything known about one fact: sources, conflicts, and uses."""
        facts = [dict(r) for r in self.conn.execute(
            "SELECT f.*, s.name AS source_name, s.path AS source_path FROM fact f"
            " LEFT JOIN source s ON s.id = f.source_id WHERE f.key=? ORDER BY f.id", (key,)).fetchall()]
        conflicts = [dict(r) for r in self.conn.execute(
            "SELECT * FROM fact_conflict WHERE key=? ORDER BY id", (key,)).fetchall()]
        uses = [dict(r) for r in self.conn.execute(
            "SELECT * FROM fill_op WHERE field=? ORDER BY id", (key,)).fetchall()]
        reviews = [dict(r) for r in self.conn.execute(
            "SELECT * FROM review_item WHERE field=? ORDER BY id", (key,)).fetchall()]
        return {"key": key, "facts": facts, "conflicts": conflicts, "fill_ops": uses,
                "reviews": reviews}

    def summary(self) -> dict[str, Any]:
        def one(sql: str) -> Any:
            row = self.conn.execute(sql).fetchone()
            return row[0] if row else 0

        by_decider = {
            r["decided_by"] or "unknown": r["n"]
            for r in self.conn.execute(
                "SELECT decided_by, COUNT(*) AS n FROM rule GROUP BY decided_by").fetchall()
        }
        by_review = {
            r["reason"] or "unknown": r["n"]
            for r in self.conn.execute(
                "SELECT reason, COUNT(*) AS n FROM review_item GROUP BY reason").fetchall()
        }
        return {
            "sessions": one("SELECT COUNT(*) FROM session"),
            "sources": one("SELECT COUNT(*) FROM source"),
            "facts": one("SELECT COUNT(*) FROM fact"),
            "facts_with_value": one("SELECT COUNT(*) FROM fact WHERE value IS NOT NULL AND value<>''"),
            "facts_missing": one("SELECT COUNT(*) FROM fact WHERE value IS NULL OR value=''"),
            "conflicts_open": one("SELECT COUNT(*) FROM fact_conflict WHERE resolved=0"),
            "templates": one("SELECT COUNT(*) FROM template"),
            "rules": one("SELECT COUNT(*) FROM rule"),
            "rules_by_decider": by_decider,
            "fill_ok": one("SELECT COUNT(*) FROM fill_op WHERE status='filled'"),
            "fill_skipped": one("SELECT COUNT(*) FROM fill_op WHERE status<>'filled'"),
            "reviews_open": one("SELECT COUNT(*) FROM review_item WHERE resolved_at IS NULL"),
            "reviews_resolved": one("SELECT COUNT(*) FROM review_item WHERE resolved_at IS NOT NULL"),
            "reviews_by_reason": by_review,
            "roles_confirmed": one("SELECT COUNT(*) FROM role_gate"),
            "roles_open_questions": one("SELECT COUNT(*) FROM role_question WHERE resolved_at IS NULL"),
            "role_gate_ok": not self.open_role_questions(),
        }
