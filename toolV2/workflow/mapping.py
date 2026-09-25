"""Expose concrete template gaps so the model maps facts, not hand-written code."""
from office_kit import doc_fill
from office_kit.common import OfficeKitError
from office_kit.store_v2 import sha256_file
from office_kit.harness import candidates_for
from office_kit.template_slots import discover_slots
from office_kit.target_validation import validate_target

SLOT_DISCOVERY_VERSION = 2


def contains(outer, inner):
    return outer[:2] == inner[:2] and outer[2] <= inner[2] and outer[3] >= inner[3]


def intersects(left, right):
    return any(a[:2] == b[:2] and (contains(a, b) or contains(b, a)
               or max(a[2], b[2]) < min(a[3], b[3])) for a in left for b in right)


def fact_catalog(store):
    from office_kit.fact_catalog import qualified_facts
    return {**store.facts_by_key(), **qualified_facts(store)}


def physical(path, target, engine=None):
    if target.get('kind') == 'multi':
        return set().union(*(physical(path, t, engine) for t in target['targets']))
    if path.suffix.lower() == '.xlsx':
        import openpyxl
        wb = openpyxl.load_workbook(path)
        try:
            sheet = target.get('sheet', 0)
            ws = wb.worksheets[sheet] if isinstance(sheet, int) else wb[sheet]
            text = str(ws[target['cell']].value or '')
            start = target.get('span_start')
            end = target.get('span_end')
            if start is None:
                start = text.find(target['anchor']) + len(target['anchor']) if target.get('anchor') else 0
                end = text.index(target['before'], start) if target.get('before') else len(text)
            return {(ws.title, target['cell'], start, end)}
        finally:
            wb.close()
    engine = engine or doc_fill.XmlEngine(path)
    if target.get('kind') == 'cell':
        cell = engine.document.tables[int(target['table'])].cell(int(target['row']), int(target['col']))
        raw = engine.xml_for(cell.paragraphs[0])
        return {('word/document.xml', raw.start, 0, len(raw.text))}
    result = set()
    for par, start, end in engine.find_anchor_hits(target):
        raw = engine.xml_for(par)
        result.add((target.get('part', 'word/document.xml'), raw.start, start, end))
    return result


def inspect(store, path, record, facts):
    """Every discovered writable position is filled, expressly blank, or exposed."""
    if (record.get('slot_hash') != sha256_file(path)
            or record.get('slot_version') != SLOT_DISCOVERY_VERSION):
        record['slots'] = discover_slots(path)
        record['slot_hash'] = sha256_file(path)
        record['slot_version'] = SLOT_DISCOVERY_VERSION
    engine = doc_fill.XmlEngine(path) if path.suffix.lower() == '.docx' else None
    covered = set()
    for rule in store.rules_for(path):
        try:
            validate_target(path, rule['target'])
            covered.update(physical(path, rule['target'], engine))
        except (OfficeKitError, ValueError, KeyError, IndexError, TypeError):
            continue
    skipped = record.get('blank_slots', {})
    missing = []
    for slot in record['slots']:
        if slot.get('protected') or slot['id'] in skipped:
            continue
        positions = physical(path, slot['target'], engine)
        if positions and all(any(contains(c, p) for c in covered) for p in positions):
            continue
        missing.append({**slot, 'candidates': candidates_for(slot.get('label', ''), facts)})
    record['unmapped_slots'] = missing
    record['coverage'] = {'detected': len(record['slots']), 'unmapped': len(missing),
                          'protected': sum(bool(s.get('protected')) for s in record['slots']),
                          'left_blank': len(skipped)}
    return missing


def visible_facts(facts):
    return [{'field': key, 'value': meta.get('value'), 'entity': meta.get('entity_name'),
             'source': meta.get('provenance'), 'fact_id': meta.get('fact_id', meta.get('id')),
             'entity_id': meta.get('entity_id'),
             'candidates': meta.get('_candidates', [])}
            for key, meta in facts.items()]


def visible_slot(slot):
    """Keep one useful excerpt, not repeated full paragraphs and fact payloads."""
    target = slot['target']
    context = dict(slot.get('context', {}))
    if 'expected_text' in target:
        raw = target['expected_text']
        start, end = target['span_start'], target['span_end']
        context['paragraph'] = raw[max(0, start-160):start] + '【填写处】' + raw[end:end+160]
    return {'id': slot['id'], 'label': slot.get('label'),
            'context': {key: value[:360] for key, value in context.items()},
            'target': {key: value for key, value in target.items() if key != 'expected_text'},
            'candidates': [{'field': c['field'], 'score': c['score']}
                           for c in slot.get('candidates', [])[:3]],
            'full_context': {'section': 'context', 'slot_id': slot['id']}}


def auto_map(store, path, record, facts, batch):
    """Automatically map only unambiguous labels; leave sentence meanings to model."""
    missing = inspect(store, path, record, facts)
    role_count = len({r['entity_id'] for r in store.roles() if r['role'] in ('借款人', '保证人')})
    added = 0
    for slot in missing:
        candidates = slot['candidates']
        if not candidates or candidates[0]['score'] < .85:
            continue
        if len(candidates) > 1 and candidates[1]['score'] == candidates[0]['score']:
            continue
        field = candidates[0]['field']
        if role_count > 1 and any(field.startswith(role) and role not in slot.get('label', '')
                                  and role not in path.stem for role in ('借款人', '保证人')):
            continue
        item = {'field': field, 'slot_id': slot['id'], 'target': {**slot['target'], 'slot_id': slot['id']}}
        target = combine(store, path, item)
        validate_target(path, target)
        tid = store.register_template(path, batch)
        store.add_rule(tid, field, slot.get('label') or field, target,
                       confidence=candidates[0]['score'], decided_by='auto', batch_no=batch)
        added += 1
    return added


def compile_updates(store, work, task, updates):
    """Validate the entire proposal before any shared rule is changed."""
    store.batch_scope = task['batch']
    facts = fact_catalog(store)
    checked, blanks = [], []
    seen = set()
    for item in updates:
        name = item['template']
        path = (work / name).resolve(strict=True)
        name = path.relative_to(work).as_posix()
        if name not in task['targets']:
            raise ValueError('只能修改当前任务的目标模板')
        if task['documents'][name].get('slot_hash') != sha256_file(path):
            raise ValueError('目标文件已变化，请先继续原任务取得最新填写位置；此次没有修改规则')
        if not item.get('slot_id'):
            validate_target(path, item['target'])
            checked.append((path, item))
            continue
        slots = {s['id']: s for s in task['documents'][name].get('slots', [])}
        slot = slots.get(item['slot_id'])
        if slot is None:
            raise ValueError(f'{name} 的空位已变化，请使用工具返回的当前 slot_id')
        identity = (name, slot['id'])
        if identity in seen:
            raise ValueError(f'{name} 的同一空位重复提交，未写入规则')
        seen.add(identity)
        if item.get('leave_blank'):
            reason = str(item.get('reason') or '').strip()
            if not reason:
                raise ValueError('留空需要说明缺失、不确定或不适用的原因')
            blanks.append((name, slot['id'], reason))
            continue
        if slot.get('protected'):
            raise ValueError('签字、签核位置不能通过填表映射自动代填')
        field = item.get('field')
        if field not in facts:
            raise ValueError(f'来源中没有字段 {field!r}，请从 available_fields 中选择或明确留空')
        target = dict(slot['target'])
        target['slot_id'] = slot['id']
        validate_target(path, target)
        checked.append((path, {**item, 'field': field, 'target': target,
                               'label': slot.get('label') or field}))
    return checked, blanks


def combine(store, path, item):
    """Add a precise position; retain other positions for the same source field."""
    field, target = item['field'], item['target']
    old = next((r for r in store.rules_for(path) if r['field'] == field), None)
    if not item.get('slot_id'):
        if old and old['target'] != target and item.get('expected_target') != old['target']:
            raise ValueError('现有位置与准备修改的位置不一致；请使用返回的 current_positions，不按旧编号修改')
        return target
    if not old:
        return target
    engine = doc_fill.XmlEngine(path) if path.suffix.lower() == '.docx' else None
    current = physical(path, target, engine)
    retained = []
    for child in old['target'].get('targets', [old['target']]):
        try:
            old_positions = physical(path, child, engine)
            if intersects(old_positions, current):
                continue
            validate_target(path, child)
            retained.append(child)
        except (OfficeKitError, ValueError, KeyError, IndexError, TypeError):
            continue
    retained.append(target)
    return {'kind': 'multi', 'targets': retained} if len(retained) > 1 else target


def retire_overlap(store, path, target, keep_field, batch):
    """Replace this exact location's prior mapping, never delete rules by IDs."""
    engine = doc_fill.XmlEngine(path) if path.suffix.lower() == '.docx' else None
    wanted = physical(path, target, engine)
    for rule in list(store.rules_for(path)):
        if rule['field'] == keep_field:
            continue
        kept, removed = [], False
        for child in rule['target'].get('targets', [rule['target']]):
            try:
                overlaps = intersects(physical(path, child, engine), wanted)
            except (OfficeKitError, KeyError, IndexError, ValueError, TypeError):
                overlaps = False
            if overlaps:
                removed = True
            else:
                kept.append(child)
        if removed:
            store.disable_rule(path, rule['field'], rule['target'], '同一明确空位已重新关联字段，保留旧规则记录')
            if kept:
                replacement = kept[0] if len(kept) == 1 else {'kind': 'multi', 'targets': kept}
                store.add_rule(rule['template_id'], rule['field'], rule['label'], replacement,
                               confidence=rule['confidence'], decided_by=rule['decided_by'], batch_no=batch)
