"""Portable template positions only; never export a customer's fact database."""
from __future__ import annotations

import json
import hashlib
import difflib
import os
import shutil
from pathlib import Path

from .common import OfficeKitError, Result, unique_path, write_text
from .store_v2 import StoreV2, sha256_file

FORMAT = "office-template-rules-v1"
CATALOG = Path(os.environ.get("OFFICE_RULE_CATALOG") or
               (Path(__file__).resolve().parents[1] / "template_rules_verified"))


def rules_digest(store: StoreV2, path: Path) -> str | None:
    """Fingerprint the exact position rules used by one fill preview."""
    if not store.rules_for(path):
        return None
    pack = build_pack(store, [path])
    raw = json.dumps(pack, ensure_ascii=False, sort_keys=True)
    return hashlib.sha256(raw.encode("utf-8")).hexdigest()


def suggest_similar(path: Path, catalog: Path = CATALOG) -> list[str]:
    """Advisory only: filename resemblance never authorizes old positions."""
    if not catalog.is_dir():
        return []
    choices = []
    digest = sha256_file(path)
    for file in catalog.glob("*.json"):
        if file.stem == digest:
            continue
        try:
            pack = json.loads(file.read_text(encoding="utf-8"))
            item = pack["templates"][0]
            name = str(item["name"])
        except (OSError, ValueError, KeyError, IndexError, TypeError):
            continue
        if Path(name).suffix.lower() != path.suffix.lower():
            continue
        score = difflib.SequenceMatcher(None, path.stem, Path(name).stem).ratio()
        if score >= 0.65:
            choices.append((score, name))
    return [name for _score, name in sorted(choices, reverse=True)[:3]]


def reuse_verified(store: StoreV2, path: Path, batch_no: str,
                   catalog: Path = CATALOG) -> dict | None:
    """Reuse only a byte-identical, previously confirmed template."""
    file = catalog / (sha256_file(path) + ".json")
    if not file.is_file():
        return None
    try:
        pack = json.loads(file.read_text(encoding="utf-8"))
    except (OSError, ValueError) as exc:
        raise OfficeKitError("已复核模板规则文件无法读取：%s" % file) from exc
    if pack.get("format") != FORMAT or len(pack.get("templates", [])) != 1:
        raise OfficeKitError("已复核模板规则文件损坏：%s" % file)
    item = pack["templates"][0]
    if item.get("sha256") != sha256_file(path):
        raise OfficeKitError("已复核模板规则与当前文件不符：%s" % path.name)
    # Renaming an otherwise identical file does not change its fill positions.
    item = {**item, "name": path.name}
    return load_pack(store, {"format": FORMAT, "templates": [item]}, [path], batch_no)


def publish_confirmed(store: StoreV2, path: Path,
                      catalog: Path = CATALOG) -> str:
    """Save positions only after this template's latest output was signed by user."""
    row = store.conn.execute(
        "SELECT a.run_id, a.status, a.opened_ok, a.signoff_decision, a.signoff_fingerprint"
        " FROM fill_op a JOIN template t ON t.id=a.template_id"
        " WHERE a.kind='artifact' AND t.path=? ORDER BY a.id DESC LIMIT 1",
        (str(path),)).fetchone()
    if not row or row["status"] != "ok" or row["opened_ok"] != 1 \
            or row["signoff_decision"] != "approve" \
            or not (row["signoff_fingerprint"] or "").startswith("confirmed-by-user|"):
        raise OfficeKitError("这份模板的最新成品尚未经使用者复核签核，不能加入自动复用目录：%s" % path.name)
    run = store.conn.execute("SELECT preflight_json FROM fill_op WHERE kind='run' AND run_id=?",
                             (row["run_id"],)).fetchone()
    preflight = json.loads(run["preflight_json"]) if run and run["preflight_json"] else {}
    snapshot = next((item for item in preflight.get("templates", [])
                     if item.get("name") == path.name and item.get("sha256") == sha256_file(path)), None)
    if not snapshot or snapshot.get("rules_sha256") != rules_digest(store, path):
        raise OfficeKitError("签核后模板或填写位置发生过变化，不能作为已复核规则自动复用：%s" % path.name)
    pack = build_pack(store, [path])
    data = json.dumps(pack, ensure_ascii=False, indent=2)
    catalog.mkdir(parents=True, exist_ok=True)
    dst = catalog / (sha256_file(path) + ".json")
    updated = False
    if dst.exists():
        old = json.loads(dst.read_text(encoding="utf-8"))
        def positions(one):
            item = one["templates"][0]
            return {r["field"]: (r["label"], r["target"], r["is_required"])
                    for r in item["rules"]}
        if positions(old) == positions(pack):
            return "already"
        history = catalog / "history"
        history.mkdir(exist_ok=True)
        # Keep the old reviewed version before making the newer signed one live.
        shutil.copy2(dst, unique_path(history / dst.name))
        updated = True
    temp = catalog / (dst.name + ".tmp")
    try:
        temp.write_text(data, encoding="utf-8")
        os.replace(temp, dst)
    finally:
        temp.unlink(missing_ok=True)
    return "updated" if updated else "published"


def build_pack(store: StoreV2, templates: list[Path]) -> dict:
    items = []
    for path in templates:
        rows = store.rules_for(path)
        if not rows:
            raise OfficeKitError("模板还没有可保存的填写规则：%s" % path.name)
        digest = sha256_file(path)
        if any(row["tpl_sha256"] != digest for row in rows):
            raise OfficeKitError("模板已变化，请重新预演后再保存规则：%s" % path.name)
        items.append({"name": path.name, "sha256": digest, "rules": [
            {"field": row["field"], "label": row["label"],
             "target": row["target"], "is_required": bool(row["is_required"]),
             "match_kind": row["match_kind"], "confidence": row["confidence"],
             "decided_by": row["decided_by"]}
            for row in rows]})
    return {"format": FORMAT, "templates": items}


def load_pack(store: StoreV2, pack: dict, templates: list[Path], batch_no: str) -> dict:
    if pack.get("format") != FORMAT or not isinstance(pack.get("templates"), list):
        raise OfficeKitError("规则包格式不对")
    by_identity = {}
    for entry in pack["templates"]:
        if not isinstance(entry, dict) or not isinstance(entry.get("rules"), list):
            raise OfficeKitError("规则包里的模板记录不完整")
        identity = (entry.get("name"), entry.get("sha256"))
        if identity in by_identity:
            raise OfficeKitError("规则包里重复登记了同一模板")
        by_identity[identity] = entry
    # Check every requested template and every rule before changing the DB.
    selected = []
    for path in templates:
        entry = by_identity.get((path.name, sha256_file(path)))
        if entry is None:
            raise OfficeKitError("规则包没有这份完全相同的模板：%s" % path.name)
        fields = set()
        for rule in entry["rules"]:
            if not isinstance(rule, dict) or not isinstance(rule.get("field"), str) \
                    or not isinstance(rule.get("target"), dict) or not rule.get("target"):
                raise OfficeKitError("规则包中 %s 的字段或位置无效" % path.name)
            if rule["field"] in fields:
                raise OfficeKitError("规则包中 %s 的字段重复：%s" % (path.name, rule["field"]))
            fields.add(rule["field"])
            if rule.get("match_kind") not in ("exact", "normalized") \
                    or rule.get("decided_by") not in ("human", "model", "auto"):
                raise OfficeKitError("规则包中 %s 的规则属性无效" % path.name)
        selected.append((path, entry))
    added = already = conflicts = 0
    for path, entry in selected:
        tid = store.register_template(path, batch_no)
        current = {row["field"]: row for row in store.rules_for(path)}
        for rule in entry["rules"]:
            old = current.get(rule["field"])
            if old:
                if old["target"] == rule["target"] and old["label"] == rule["label"]:
                    already += 1
                else:
                    conflicts += 1  # Never overwrite work done in this workspace.
                continue
            store.add_rule(tid, rule["field"], rule.get("label"), rule["target"],
                           is_required=int(bool(rule.get("is_required"))),
                           match_kind=rule["match_kind"], confidence=rule.get("confidence"),
                           decided_by=rule["decided_by"], batch_no=batch_no)
            added += 1
    return {"templates": len(selected), "added": added,
            "already_present": already, "conflicts_skipped": conflicts}


def cmd_db_rules_export(args) -> Result:
    from .harness import _open

    res = Result("db-rules-export")
    _wr, store = _open(args)
    try:
        templates = [Path(p) for p in args.input]
        pack = build_pack(store, templates)
        dst = write_text(unique_path(Path(args.out)),
                         json.dumps(pack, ensure_ascii=False, indent=2))
        res.add_artifact(dst, "可复用的模板位置规则；不含客户事实")
        res.data.update({"templates": len(pack["templates"]),
                         "rules": sum(len(item["rules"]) for item in pack["templates"]),
                         "file": str(dst)})
        return res
    finally:
        store.close()


def cmd_db_rules_import(args) -> Result:
    from .harness import _open

    res = Result("db-rules-import")
    _wr, store = _open(args)
    try:
        pack = json.loads(Path(args.pack).read_text(encoding="utf-8"))
        batch_no = args.batch or store.current_batch()
        if not batch_no:
            raise OfficeKitError("先导入源材料并建立批次，再导入模板规则")
        res.data.update(load_pack(store, pack, [Path(p) for p in args.input], batch_no))
        if res.data["conflicts_skipped"]:
            res.warn("已有 %d 条不同的规则，未覆盖；请核对后再决定。"
                     % res.data["conflicts_skipped"])
        return res
    finally:
        store.close()
