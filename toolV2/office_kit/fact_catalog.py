"""Build explicit subject aliases for the fill planner.

The catalog is a read-only view.  It deliberately keeps the original fact
metadata and adds only the relationship metadata needed by the planner.
"""
from __future__ import annotations

import re
import hashlib
from collections import defaultdict


_GUARANTOR_MARK = re.compile(r"保证人\s*(\d+)")
_ROLE_FIELDS = {"借款人", "保证人", "法定代表人"}
_LEGAL_KEYS = {"借款人法定代表人", "保证人法定代表人", "法定代表人"}
_ID_FIELDS = {"证件类型", "证件号码", "身份证号码", "身份证号", "护照号码", "护照号"}
_TARGET_ROLE_RE = re.compile(r"(?:保证人|担保人)[（(]?\s*(\d+)\s*[）)]?|借款人|授信申请人")
# A form can contain a repeated signer block which its own text says to add or
# remove as signers change.  That instruction is evidence that the next blank
# is optional, not evidence that it belongs to the borrower in an earlier block.
_OPTIONAL_PARTY_RE = re.compile(
    r"(?:可\s*)?(?:根据|按)\s*(?:签署|签约|签字|参与)(?:方|主体).{0,12}"
    r"(?:个数|数量).{0,12}(?:增减|增删|调整)|"
    r"(?:签署|签约|签字)(?:方|主体).{0,12}(?:个数|数量).{0,12}(?:增减|增删|调整)"
)
_COMPANY_BLOCK_HEADER_RE = re.compile(
    r"(?:企业|公司|单位)\s*(?:名称|名)?\s*(?:[（(][^）)\r\n]*[）)]\s*)?[：:]")
_BLOCK_BOUNDARY_RE = re.compile(
    r"经过对.*(?:签约|签章|签署).*核实|(?:签约|签章|签署)核实(?:书|表|意见)?|"
    r"^\s*(?:第?[一二三四五六七八九十]+[、.]|\d+[、.])"
)


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
    # A repeated signer block may span several identity fields (company name,
    # representative, address, contact and bank account).  Find its nearest
    # company-name header instead of relying on a paragraph count.  An empty
    # paragraph or a new verification conclusion/section ends that block, so
    # the optional instruction cannot leak into the next normal borrower area.
    tier_texts = [" ".join(tier).strip() for tier in tiers]
    header_index = next((i for i, text in enumerate(tier_texts)
                         if _COMPANY_BLOCK_HEADER_RE.search(text)), None)
    additional_party = False
    additional_party_block_id = None
    binding_text = ""
    if header_index is not None:
        header = tier_texts[header_index]
        directive_index = header_index + 1
        while directive_index < len(tier_texts) and tier_texts[directive_index] == header:
            directive_index += 1
        directive = (tier_texts[directive_index]
                     if directive_index < len(tier_texts) else "")
        # These are the fields after the header and before the current target.
        # Any semantic divider makes the older header irrelevant.
        later_block = tier_texts[:header_index]
        bounded = not any(not text or _BLOCK_BOUNDARY_RE.search(text) for text in later_block)
        binding_text = " ".join(tier_texts[:header_index + 1] + [directive])
        explicit_role, _explicit_number = mark([binding_text])
        entity_names = _entity_names(store)
        explicit_entity = any(name and name in binding_text
                              for names in entity_names.values() for name in names)
        additional_party = bool(bounded and _OPTIONAL_PARTY_RE.search(header + " " + directive)
                                and explicit_role is None and not explicit_entity)
        if additional_party:
            # The header's physical location, rather than its business text,
            # separates repeated optional signer blocks in one template.
            offset = max(header_index - 1, 0)
            if target.get('kind') == 'anchor':
                location = (target.get('part', 'word/document.xml'),
                            int(target.get('paragraph_index', 0)) - offset)
            elif target.get('kind') == 'cell':
                location = ('table', int(target.get('table', 0)),
                            int(target.get('row', 0)) - offset)
            else:
                location = (target.get('sheet'), target.get('cell'), header_index)
            identity = repr((str(path.resolve()), target.get('kind'), location)).encode('utf-8')
            additional_party_block_id = 'optional-' + hashlib.sha256(identity).hexdigest()[:20]
    else:
        explicit_role, _explicit_number = mark([" ".join(tier) for tier in tiers[:3]])

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
    guaranteed_credit_amount = bool(
        re.search(r"为[^。；;]*公司[^。；;]*申请(?:的)?[^。；;]*$", raw[:start]) and
        re.match(r"\s*(?:亿元|万元|元)\s*授信", raw[end:]) and
        re.search(r"保证|担保", raw[end:]))
    if guarantee_recipient or guaranteed_credit_amount:
        role, number = "借款人", None
    claims = _roles(store)
    if additional_party:
        # Do not let a role found in an older preceding section default this
        # optional signer block to the borrower.  A human/model may still map
        # it explicitly; this only prevents automatic inference.
        role, number = None, None
    filename = path.stem
    if role is None and not additional_party:
        named = [r for r in ("借款人", "保证人") if r in filename]
        role = named[0] if len(named) == 1 else None
    declared_role, declared_number = role, number
    candidates = [eid for eid, items in claims.items()
                  if any(item["role"] == role for item in items)] if role else []
    if role is None and not additional_party:
        borrowers = [eid for eid, items in claims.items()
                     if any(item["role"] == "借款人" for item in items)]
        # A unique borrower is not a license to fill an otherwise anonymous
        # repeated “法定代表人/联系电话” slot.  Default only when the filename
        # or actual nearby text declares ordinary borrower business context.
        business_text = " ".join(" ".join(tier) for tier in tiers)
        borrower_context = bool(re.search(r"借款|授信|用信|贷款|客户|本公司|申请|签约核实|签章核实", filename + " " + business_text))
        if len(borrowers) == 1 and borrower_context:
            role, candidates = "借款人", borrowers
    if role == "保证人" and number is not None:
        numbered = [eid for eid in candidates if any(
            _GUARANTOR_MARK.search(item["evidence"] or "") and
            int(_GUARANTOR_MARK.search(item["evidence"] or "").group(1)) == number
            for item in claims[eid])]
        candidates = numbered
    if role == '保证人' and number is None and any(word in path.stem for word in ('法定代表人身份证明', '董事会决议')):
        companies = [eid for eid in candidates if store.is_company_evidenced(eid)]
        if companies:
            candidates = companies
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
        "法定代表人证件号码" if re.search(r"(?:法定代表人|法人代表)(?:或授权代理人)?(?:身份证|证件)(?:号|号码)[：:]?\s*$", raw[:start]) else
        "法定代表人" if re.search(r"(?:法定代表人|法人代表)(?:或授权代理人)?(?:姓名)?\s*[：:]?\s*$", before) else
        "名称" if re.search(r"(?:保证人|担保人).{0,4}(?:名称|姓名)", slot_text) else None)
    if re.search(r'配偶', before):
        field_hint = '配偶姓名及证件号码'
    elif certificate_template and re.match(r'\s*职务', after):
        field_hint = '职务'
    elif guaranteed_credit_amount:
        field_hint = '授信金额'
    elif guarantee_recipient:
        field_hint = '名称'
    elif (re.search(r'经过对\s*$', before) and re.match(r'\s*签约', after)) or re.match(r'\s*[（(]以下简称[“"\s]*(?:借款人|保证人)', after):
        field_hint = '名称'
    elif re.match(r'\s*(?:有限公司|有限责任公司|公司|董事会)', after):
        field_hint = '名称'
    elif event_place := re.search(r'((?:会议|核实|签约|签章|签署)地点)[：:]?\s*$', before):
        field_hint = event_place.group(1)
    elif re.search(r'(?:业务品种|用信品种|单笔用信业务)[：:]?\s*$', before):
        field_hint = '用信业务品种'
    elif re.search(r'(?:地址|住所|办公地点)[：:]?\s*$', before):
        field_hint = '地址'
    elif re.search(r'(?:本次(?:申请)?(?:用信|借款)|放款|本次业务金额)(?:人民币)?\s*$', before):
        field_hint = '借款金额'
    elif re.search(r'合同金额\s*$', before):
        if re.search(r'流动资金贷款合同|借款合同',raw[:start]):
            field_hint = '借款金额'
        elif '授信' in raw[:start]:
            field_hint = '授信金额'
    elif re.search(r'(?:合同编号|编号为)[^，。；;（）()]{0,40}$', raw[:start]):
        field_hint = '合同编号'
    elif re.search(r"开户(?:银)?行及账号\s*[：:]?\s*$", before):
        field_hint = '开户行及账号'
    elif re.search(r"(?:姓名|名称)及(?:身份证|证件)(?:号|号码)?\s*[：:]?\s*$", before):
        field_hint = ('法定代表人姓名及证件号码' if '法定代表人' in raw[:start]
                      else '名称及证件号码')
    reason = ("可选签署方重复区块未声明角色、编号或主体名称；未默认复制借款人"
              if additional_party else
              ("位置文字明确%s%s" % (role, number if number is not None else "")
               if role else "位置正文和文件名均未声明主体"))
    return {"role": role, "number": number, "entity_ids": sorted(candidates),
            "field_hint": field_hint,
            "additional_party": additional_party,
            "declared_role": declared_role, "declared_number": declared_number,
            "additional_party_block_id": additional_party_block_id,
            "reason": reason}


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
    _available_credit(out)
    return out


def _available_credit(out):
    """Expose credit minus used credit with both source amounts as evidence."""
    from decimal import Decimal
    from .value_fit import _amount_unit, _converted_amount
    from .common import OfficeKitError
    prefixes = {m.group(0) for key in out if (m := re.match(r'^(?:借款人|保证人\d*)', key))}
    for prefix in prefixes:
        if any(prefix + key in out for key in ('可用授信额度', '可用额度', '可使用授信额度')):
            continue
        credit = next((out[prefix+k] for k in ('授信金额','授信额度') if prefix+k in out), None)
        used = next((out[prefix+k] for k in ('已用额度','已使用额度','已使用授信额度') if prefix+k in out), None)
        if not credit or not used or credit.get('_ambiguous') or used.get('_ambiguous'):
            continue
        if (credit.get('entity_id'),credit.get('_relation_owner_eid')) != (used.get('entity_id'),used.get('_relation_owner_eid')):
            continue
        unit = _amount_unit(str(credit.get('value') or ''))
        if not unit or not _amount_unit(str(used.get('value') or '')):
            continue
        try:
            available = Decimal(_converted_amount(str(credit['value']),unit)) - Decimal(_converted_amount(str(used['value']),unit))
        except OfficeKitError:
            continue
        if available < 0:
            continue
        meta = {k:v for k,v in credit.items() if k not in ('id','fact_id')}
        meta.update(value=format(available.normalize(),'f')+unit, source_kind='computed',
                    provenance=f"可用额度=授信额度({credit['value']})-已用额度({used['value']})；"+
                               str(credit.get('provenance') or '')+'；'+str(used.get('provenance') or ''),
                    fact_ids=[p.get('fact_id',p.get('id')) for p in (credit,used)],
                    _base_field='可用授信额度',_qualified=True)
        out[prefix+'可用授信额度']=meta


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
