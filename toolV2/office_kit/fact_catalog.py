"""Build explicit subject aliases for the fill planner.

The catalog is a read-only view.  It deliberately keeps the original fact
metadata and adds only the relationship metadata needed by the planner.
"""
from __future__ import annotations

import re
from collections import defaultdict


_GUARANTOR_MARK = re.compile(r"保证人\s*(\d+)")
_ROLE_FIELDS = {"借款人", "保证人", "法定代表人"}
_LEGAL_KEYS = {"借款人法定代表人", "保证人法定代表人", "法定代表人"}
_ID_FIELDS = {"证件类型", "证件号码", "身份证号码", "身份证号", "护照号码", "护照号"}


def _scope(store):
    return getattr(store, "batch_scope", None)


def _rows(store):
    scope = _scope(store)
    sql = ("SELECT f.* FROM fact f WHERE f.superseded_by IS NULL "
           "AND (? IS NULL OR f.batch_no=?) ORDER BY f.id")
    return [dict(r) for r in store.conn.execute(sql, (scope, scope))]


def _entity_names(store):
    rows = store.conn.execute("SELECT entity_id,name FROM entity_name ORDER BY id")
    names = defaultdict(list)
    for row in rows:
        names[int(row[0])].append(str(row[1]))
    return names


def _roles(store):
    scope = _scope(store)
    sql = ("SELECT cr.entity_id,cr.role,cr.evidence FROM case_role cr "
           "JOIN \"case\" c ON c.id=cr.case_id "
           "WHERE (? IS NULL OR c.batch_no=?) ORDER BY cr.id")
    out = defaultdict(list)
    for row in store.conn.execute(sql, (scope, scope)):
        out[int(row[0])].append({"role": row[1], "evidence": row[2] or ""})
    return out


def _role_number(evidence: str) -> int | None:
    match = _GUARANTOR_MARK.search(evidence or "")
    return int(match.group(1)) if match else None


def _base_field(key: str) -> str:
    for role in ("借款人", "保证人"):
        if key.startswith(role):
            return key[len(role):] or key
    return key


def _meta(row: dict, *, base: str, owner: int | None, entity_id: int | None = None,
          entity_name: str | None = None) -> dict:
    value = dict(row)
    if "id" in row:
        value.setdefault("fact_id", row["id"])
    if entity_id is not None:
        value["entity_id"] = entity_id
    if entity_name is not None:
        value["entity_name"] = entity_name
    value.update({"_base_field": base, "_relation_owner_eid": owner,
                  "_qualified": True})
    return value


def _add(out: dict, key: str, value: dict) -> None:
    """Add an alias; collisions remain explicit ambiguity."""
    old = out.get(key)
    if old is None:
        out[key] = value
        return
    if ((not old.get("_qualified")) and value.get("_qualified") and
            old.get("entity_id") == value.get("entity_id") and
            old.get("value") == value.get("value")):
        out[key] = value
        return
    if old.get("_ambiguous"):
        candidates = old["_candidates"]
    else:
        candidates = [old]
    if not any((c.get("entity_id"), c.get("value")) ==
               (value.get("entity_id"), value.get("value")) for c in candidates):
        candidates.append(value)
    if len(candidates) == 1:
        out[key] = candidates[0]
        return
    out[key] = {"key": key, "value": None, "entity_id": None,
                "_ambiguous": True, "_candidates": candidates,
                "_base_field": value.get("_base_field"),
                "_relation_owner_eid": None, "_qualified": True}


def _person_for_name(store, names, value: str, person_ids: set[int]) -> int | None:
    if not value:
        return None
    exact = [eid for eid in person_ids if value in names.get(eid, [])]
    if len(exact) == 1:
        return exact[0]
    # Some imported databases have the name only in a fact.  Restrict this
    # fallback to an exact current value.
    rows = store.conn.execute(
        "SELECT DISTINCT f.entity_id FROM fact f "
        "WHERE f.entity_id IS NOT NULL AND f.superseded_by IS NULL "
        "AND f.value=?", (value,))
    exact = [int(r[0]) for r in rows if int(r[0]) in person_ids]
    return exact[0] if len(exact) == 1 else None


def qualified_facts(store) -> dict[str, dict]:
    """Return current-batch facts plus explicit subject-qualified aliases.

    No rows are inserted or updated.  Ambiguous aliases are represented with
    ``_ambiguous`` and ``_candidates`` just like ``facts_by_key``.
    """
    out = dict(store.facts_by_key())
    rows = _rows(store)
    names = _entity_names(store)
    roles = _roles(store)
    # Older migrated databases may classify every entity as legal_person.
    # Exact unique name/value evidence is the association; entity_type alone
    # must not discard the representative.
    person_ids = {int(r[0]) for r in store.conn.execute("SELECT id FROM entity")}

    guarantors = defaultdict(list)
    for eid, claims in roles.items():
        for claim in claims:
            if claim["role"] == "保证人":
                guarantors[eid].append(_role_number(claim["evidence"]))

    # Ordinary role-qualified facts.  A number is accepted only from stored
    # evidence; iteration order is never used to invent 保证人1/2.
    for row in rows:
        eid = row.get("entity_id")
        if eid is None or eid not in roles:
            continue
        base = _base_field(str(row["key"]))
        if row["key"] in _LEGAL_KEYS:
            continue
        for claim in roles[eid]:
            role = claim["role"]
            if role == "法定代表人":
                continue
            if role == "保证人":
                nums = [n for n in guarantors[eid] if n is not None]
                prefix = f"保证人{nums[0]}" if len(set(nums)) == 1 else "保证人"
            else:
                prefix = role
            _add(out, prefix + base, _meta(row, base=base, owner=eid,
                                           entity_id=eid,
                                           entity_name=names.get(eid, [None])[0]))

    # Resolve company -> legal representative person from the explicit
    # company fact, then expose the person's name and identity facts.
    legal_companies = []
    for row in rows:
        if row["key"] in _LEGAL_KEYS and row.get("entity_id") is not None:
            relation = row["key"] if row["key"] in ("借款人法定代表人", "保证人法定代表人") else None
            legal_companies.append((row, relation))
    for company_fact, relation in legal_companies:
        owner = int(company_fact["entity_id"])
        role_prefix = (relation.removesuffix("法定代表人") if relation else
                       next((c["role"] for c in roles.get(owner, [])
                             if c["role"] in ("借款人", "保证人")), None))
        if role_prefix is None:
            continue
        person = _person_for_name(store, names, str(company_fact.get("value") or ""), person_ids)
        if person is None:
            continue
        rep_number = None
        if role_prefix == "保证人":
            nums = [n for n in guarantors.get(owner, []) if n is not None]
            if len(set(nums)) == 1:
                rep_number = nums[0]
        name_key = ((f"保证人{rep_number}" if rep_number is not None else "保证人")
                    + "法定代表人" if role_prefix == "保证人"
                    else (relation or "借款人法定代表人"))
        person_meta = _meta(company_fact, base=company_fact["key"], owner=owner,
                            entity_id=person, entity_name=(names.get(person) or [None])[0])
        # The company relation fact uses the same human-facing key.  Replace
        # that unqualified raw view with the explicit person association;
        # the original row remains available under its source key elsewhere.
        if name_key in out and not out[name_key].get("_qualified"):
            out[name_key] = person_meta
        else:
            _add(out, name_key, person_meta)
        for row in rows:
            if int(row.get("entity_id") or -1) != person or row["key"] not in _ID_FIELDS:
                continue
            _add(out, name_key + str(row["key"]),
                 _meta(row, base=row["key"], owner=owner, entity_id=person,
                       entity_name=(names.get(person) or [None])[0]))
    return out
