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
_TARGET_ROLE_RE = re.compile(r"(?:保证人|担保人)[（(]?\s*(\d+)\s*[）)]?|借款人|授信申请人")


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


def target_subject_context(store, path, target: dict, *, label: str = "",
                           context: dict | None = None) -> dict:
    """Return the source-backed subject scope for one target position.

    The result is intentionally small and read-only, for mapping's automatic
    rule choice: ``entity_ids`` is empty when the target has no explicit
    subject clue.  Numbered guarantors are resolved from stored role evidence,
    never from entity ordering.
    """
    # Never trust a persisted rule label: it may have been shared by a former
    # multi-target rule and describe a different table section.  Inspect the
    # actual target text first, then nearest table/paragraph context, then the
    # filename.  Each tier is evaluated separately so an earlier guarantor
    # heading cannot override an explicit “授信申请人” at this position.
    tiers: list[list[str]] = []

    def mark(values: list[str]) -> tuple[str | None, int | None]:
        matches = list(_TARGET_ROLE_RE.finditer(" ".join(values)))
        if not matches:
            return None, None
        hit = matches[-1]
        role = "保证人" if ("保证" in hit.group(0) or "担保" in hit.group(0)) else "借款人"
        return role, int(hit.group(1)) if hit.group(1) else None

    if target.get("expected_text"):
        tiers.append([str(target["expected_text"])])
    try:
        if path.suffix.lower() == ".docx" and target.get("kind") == "cell":
            from .target_validation import read_word_template
            table = read_word_template(path).document.tables[int(target["table"])]
            row, col = int(target["row"]), int(target["col"])
            tiers.append([table.cell(row, col).text] +
                         [cell.text for cell in table.rows[row].cells[:col]])
            for r in range(row - 1, max(-1, row - 8), -1):
                tiers.append([cell.text for cell in table.rows[r].cells])
        elif path.suffix.lower() == ".docx" and target.get("kind") == "anchor":
            from .target_validation import read_word_template
            engine = read_word_template(path)
            part, wanted = target.get("part", "word/document.xml"), target.get("paragraph_index")
            if wanted is not None:
                entries = [par for _element, par in engine._entries
                           if engine._part_by_el.get(par._element) == part]
                index = int(wanted)
                if 0 <= index < len(entries):
                    tiers.append([(engine.xml_for(entries[index]).text
                                   if engine.xml_for(entries[index]) else entries[index].text)])
                    def table_parent(par):
                        node = par._element
                        while node is not None:
                            if str(node.tag).endswith("}tbl"):
                                return node
                            node = node.getparent()
                        return None
                    same_table = table_parent(entries[index])
                    prior = [i for i in range(index - 1, -1, -1)
                             if same_table is not None and table_parent(entries[i]) is same_table]
                    # A table is a semantic block: its heading may be farther
                    # than 32 XML paragraphs but is safer than an arbitrary
                    # preceding section. Outside tables retain a bounded scan.
                    if not prior:
                        prior = list(range(index - 1, max(-1, index - 64), -1))
                    for i in prior:
                        par = entries[i]
                        tiers.append([(engine.xml_for(par).text if engine.xml_for(par) else par.text)])
        elif path.suffix.lower() == ".xlsx" and target.get("kind") == "xlsx_cell":
            from .target_validation import read_excel_template
            wb = read_excel_template(path)
            ws = wb[str(target.get("sheet"))] if target.get("sheet") is not None else wb.worksheets[0]
            cell = ws[str(target["cell"])]
            label_cell = target.get("label_cell")
            tiers.append([str(cell.value or ""), str(ws[str(label_cell)].value or "") if label_cell else ""] +
                         [str(ws.cell(cell.row, c).value or "") for c in range(1, cell.column)])
            for r in range(cell.row - 1, max(0, cell.row - 8), -1):
                tiers.append([str(ws.cell(r, c).value or "") for c in range(1, ws.max_column + 1)])
    except (IndexError, KeyError, TypeError, ValueError, OSError):
        pass
    role = number = None
    for tier in tiers:
        found_role, found_number = mark(tier)
        if role is None and found_role:
            role, number = found_role, found_number
            if role != "保证人" or number is not None:
                break
            # Generic “担保人法定代表人” needs the nearest preceding
            # numbered guarantor heading; do not stop at the generic label.
            continue
        if role == "保证人" and found_role == "保证人" and found_number is not None:
            number = found_number
            break
    # A guarantee document also names the recipient of its guarantee. Resolve
    # the noun at this blank, rather than the last role word in its sentence.
    raw = str(target.get("expected_text") or "")
    start, end = int(target.get("span_start", 0)), int(target.get("span_end", 0))
    guarantee_recipient = bool(
        re.search(r"为\s*$", raw[:start]) and
        re.match(r"\s*(?:有限)?公司", raw[end:]) and
        re.search(r"授信|借款|贷款", raw[end:]) and
        re.search(r"保证|担保", raw[end:]))
    if guarantee_recipient:
        role, number = "借款人", None
    claims = _roles(store)
    if role is None:
        filename = path.stem
        named = [r for r in ("借款人", "保证人") if r in filename]
        role = named[0] if len(named) == 1 else None
    candidates = [eid for eid, items in claims.items()
                  if any(item["role"] == role for item in items)] if role else []
    if role is None:
        borrowers = [eid for eid, items in claims.items()
                     if any(item["role"] == "借款人" for item in items)]
        # A unique borrower is not a license to fill an otherwise anonymous
        # repeated “法定代表人/联系电话” slot.  Default only when the filename
        # or actual nearby text declares ordinary borrower business context.
        business_text = " ".join(" ".join(tier) for tier in tiers)
        borrower_context = bool(re.search(r"借款|授信|用信|贷款|客户|本公司|申请", filename + " " + business_text))
        if len(borrowers) == 1 and borrower_context:
            role, candidates = "借款人", borrowers
    if role == "保证人" and number is not None:
        numbered = [eid for eid in candidates if any(
            _GUARANTOR_MARK.search(item["evidence"] or "") and
            int(_GUARANTOR_MARK.search(item["evidence"] or "").group(1)) == number
            for item in claims[eid])]
        candidates = numbered
    # Meaning belongs to this *span*, never to another blank in the same
    # paragraph. A certificate paragraph can contain name, ID, sex and age;
    # only the first two have source-backed relation fields.
    raw = str(target.get("expected_text") or "")
    try:
        start, end = int(target.get("span_start", 0)), int(target.get("span_end", 0))
    except (TypeError, ValueError):
        start = end = 0
    before, after = raw[max(0, start - 20):start], raw[end:end + 20]
    slot_text = before + "【填写处】" + after
    certificate_template = "法定代表人身份证明" in path.stem
    field_hint = (
        # “____同志（…证件号：____）” has two slots in one paragraph. The
        # name slot is identified by its immediate right word, before looking
        # for any later certificate text.
        "法定代表人" if certificate_template and re.match(r"\s*同志", after) else
        "法定代表人证件号码" if certificate_template and re.search(r"身份证|护照|通行证|证件号", before) else
        "法定代表人" if re.search(r"法定代表人|法人代表", slot_text) else
        "名称" if re.search(r"(?:保证人|担保人).{0,4}(?:名称|姓名)", slot_text) else None)
    if guarantee_recipient:
        field_hint = '名称'
    elif re.search(r"开户(?:银)?行及账号\s*[：:]?\s*$", before):
        field_hint = '开户行及账号'
    elif re.search(r"(?:姓名|名称)及(?:身份证|证件)(?:号|号码)?\s*[：:]?\s*$", before):
        field_hint = ('法定代表人姓名及证件号码' if '法定代表人' in raw[:start]
                      else '名称及证件号码')
    return {"role": role, "number": number, "entity_ids": sorted(candidates),
            "field_hint": field_hint,
            "reason": ("位置文字明确%s%s" % (role, number if number is not None else "")
                       if role else "位置正文和文件名均未声明主体")}


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
    _composite_facts(out)
    return out


def _composite_facts(out):
    """Join explicitly requested components from one owner, retaining evidence."""
    prefixes = {m.group(0) for key in out if (m := re.match(r'^(?:借款人|保证人\d*)', key))}
    pairs = [('开户行及账号', '开户行', '收款账号'),
             ('名称及证件号码', '名称', '证件号码'),
             ('法定代表人姓名及证件号码', '法定代表人', '法定代表人证件号码')]
    for prefix in prefixes:
        for suffix, left, right in pairs:
            # A company's representative is not the individual guarantor. A
            # printed personal-name/ID line stays empty for a company owner.
            if suffix == '名称及证件号码' and prefix + '法定代表人' in out:
                continue
            parts = [out.get(prefix + left), out.get(prefix + right)]
            if any(p and p.get('_ambiguous') for p in parts):
                continue
            parts = [p for p in parts if p and p.get('value') not in (None, '')]
            if not parts:
                continue
            a = parts[0]
            if any((a.get('entity_id'), a.get('_relation_owner_eid')) != (b.get('entity_id'), b.get('_relation_owner_eid')) for b in parts[1:]):
                continue
            if prefix + suffix in out:
                continue
            meta = {k: v for k, v in a.items() if k not in ('id', 'fact_id')}
            meta.update(value=' '.join(str(p['value']) for p in parts), source_kind='computed',
                        provenance='；'.join(str(p.get('provenance') or '') for p in parts),
                        fact_ids=[p.get('fact_id', p.get('id')) for p in parts],
                        _base_field=suffix, _qualified=True)
            out[prefix + suffix] = meta
