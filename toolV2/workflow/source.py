"""Read each source version once; retain ownership answers across task restarts."""
from office_kit import absorb
from office_kit.store_v2 import sha256_file
from .storage import digest, get, put, inside
from . import source_conflicts

# Increment when source interpretation changes, so saved tasks do not silently
# keep an older parser's ownership suggestions. Answer identities stay separate.
# v5: numbered guarantor sections retain their own current owner instead of
# letting a later “保证人2” overwrite the fields following “保证人1”.
SOURCE_PARSE_VERSION = 5
# Parser refresh must not erase unchanged user decisions. Evidence/ownership
# changes already produce different identities; retain the prior answer schema.
SOURCE_ANSWER_VERSION = 3


def prepare(store, work, task):
    signatures = [(p, sha256_file(inside(work, p))) for p in task['source']]
    signature = digest([SOURCE_PARSE_VERSION, signatures])
    task['source_content_signature'] = digest(signatures)
    task['equivalent_source_signatures'] = [digest([version, signatures])
                                            for version in range(1, SOURCE_PARSE_VERSION + 1)]
    if task.get('source_signature') == signature and task.get('source_ready'):
        return []
    rows = []
    diagnostics = []
    for name, fingerprint in signatures:
        cache_id = 'source:' + digest([SOURCE_PARSE_VERSION, name, fingerprint])
        cached = get(store.conn, 'office_v2_cache', cache_id)
        if cached is None:
            parsed, issues = absorb.absorb_with_diagnostics(inside(work, name))
            cached = {'rows': parsed, 'issues': issues}
            put(store.conn, 'office_v2_cache', cache_id, cached)
            task['counts']['source_reads'] += 1
        diagnostics.extend({'stage': 'source', 'source': name, **issue}
                           for issue in cached['issues'])
        for row in cached['rows']:
            rows.append({**row, 'n': len(rows) + 1, '_source': name})
    task['source_issues'] = diagnostics
    task['issues'] = list(diagnostics)
    task['rows'] = rows
    if not rows:
        task['source_ready'] = False
        task['issues'] = diagnostics or [{'stage': 'source', 'reason': '源文件未识别出键值，请补充材料或针对源解析修复；未猜测填写。'}]
        task['status'] = 'needs_mapping'
        return []
    task['rows'] = rows
    task['source_signature'] = signature
    task['source_ready'] = False
    question_groups = {}
    for question in absorb.source_questions(rows):
        row = next(r for r in rows if question['id'] == f"source-{r['n']}")
        linked = [r for r in rows if r['_source'] == row['_source']
                  and r.get('ownership_group') == row['ownership_group']] if row.get('ownership_group') else [row]
        field_values = sorted({(r['key'], r['value']) for r in linked})
        identity = digest([SOURCE_ANSWER_VERSION, task['batch'], row['_source'], field_values,
                           row.get('entity_name'), row.get('quote'), row.get('context')])
        question['id'] = 'source-' + identity
        for item in linked:
            item['_answer_key'] = question['id']
        # `header` and the source line number only identify one occurrence.  They
        # legitimately differ for identical facts, so they must not prevent one
        # user decision from covering every occurrence.  Options and the rest of
        # the question must still agree: a hash collision must never silently turn
        # two differently-worded decisions into one answer.
        comparable = {key: value for key, value in question.items()
                      if key not in ('id', 'header')}
        comparable['question'] = comparable['question'].replace(
            f"材料第{row.get('line')}行", '材料第<来源位置>行', 1)
        group = question_groups.get(question['id'])
        if group is None:
            question_groups[question['id']] = {
                'question': question,
                'comparable': comparable,
                'rows': linked,
            }
        elif group['comparable'] != comparable:
            raise ValueError(
                f"来源确认问题 ID 冲突：{question['id']} 对应的问题内容或选项不一致；"
                '未合并，也未保存任何回答。')
        else:
            group['rows'].extend(item for item in linked if item not in group['rows'])

    questions = []
    for question_id, group in question_groups.items():
        reply = get(store.conn, 'office_v2_cache', question_id)
        if reply is None:
            occurrences = {(r['_source'], r.get('line'), r.get('quote')) for r in group['rows']}
            if len(occurrences) > 1:
                group['question']['question'] += (
                    f" 相同信息共出现 {len(occurrences)} 次，本次选择用于这些相同条目。")
            questions.append(group['question'])
        else:
            for row in group['rows']:
                apply_one(row, reply)
    return questions if questions else source_conflicts.prepare(store, task)


def apply_one(row, answer):
    custom = answer.get('custom', '').strip()
    selected = answer.get('selected') or []
    pick = custom or selected[0]
    if not custom and pick == '暂不收录':
        row['dropped'] = True
    else:
        row['entity_name'] = pick.removeprefix('归属：').removesuffix('（推荐）').strip()
        row['assigned_by_user'] = True
        row['assumed'] = False
        row['needs'] = '收'


def save_answers(store, task, answers):
    for answer in answers:
        if answer['id'] in task.get('source_conflict_groups', {}):
            source_conflicts.save_answer(store, task, answer)
            continue
        rows = [r for r in task['rows'] if r.get('_answer_key') == answer['id']]
        if not rows:
            raise ValueError(f"回答不属于当前来源确认问题：{answer.get('id')}")
        for row in rows:
            apply_one(row, answer)
        put(store.conn, 'office_v2_cache', answer['id'], answer)
    store.event('source_review_confirmed', batch_no=task['batch'], target=task['id'],
                payload={'answers': answers}, actor='user')


def ingest(store, work, task):
    if source_conflicts.conflicting_groups(task['rows']):
        raise ValueError('来源值存在冲突，请先完成工具内的冲突确认；未开始写入事实。')
    source_ids = {}
    for name in task['source']:
        path = inside(work, name)
        source_ids[name] = store.register_source(path, task['batch'],
            copy_root=work / 'in' / sha256_file(path)[:12])[0]
    seen = set()
    for row in task['rows']:
        if row.get('dropped'):
            continue
        eid = store.ensure_entity(row['entity_name'])[0] if row.get('entity_name') else None
        key = (row['key'], eid)
        if key in seen:
            previous = [r for r in task['rows'] if r is not row and not r.get('dropped')
                        and r['key'] == row['key'] and r.get('entity_name') == row.get('entity_name')]
            if any(r['value'] != row['value'] for r in previous):
                raise ValueError(f"同一主体的 {row['key']} 在源文件中有不同值，请明确使用哪个；没有覆盖旧值。")
            continue
        seen.add(key)
        sid = source_ids[row['_source']]
        label = row.get('label', '')
        role = row.get('role') or ('借款人' if label in ('借款人', '借款人名称') else '保证人' if label in ('保证人', '保证人名称') else None)
        if eid is not None and role:
            store.set_role(eid, role, evidence=row.get('quote', ''), decided_by='source',
                           source_id=sid, batch_no=task['batch'])
        store.put_fact(task['batch'], row['key'], row['value'], entity_id=eid, source_id=sid,
                       source_kind='source', provenance=row.get('quote'), confidence=row.get('confidence'),
                       on_conflict='supersede')
    # Values removed from this source version must not survive as current facts.
    source_names = [str(inside(work, p)) for p in task['source']]
    for fact in list(store.conn.execute('SELECT f.* FROM fact f JOIN source s ON s.id=f.source_id '
                                       'WHERE f.batch_no=? AND f.superseded_by IS NULL AND f.source_kind=?',
                                       (task['batch'], 'source'))):
        original = store.conn.execute('SELECT path FROM source WHERE id=?', (fact['source_id'],)).fetchone()
        if original and original[0] in source_names and (fact['key'], fact['entity_id']) not in seen:
            store.put_fact(task['batch'], fact['key'], None, entity_id=fact['entity_id'],
                           source_id=fact['source_id'], source_kind='source', status='missing',
                           provenance='本次材料中已无此项', on_conflict='supersede')
    task['source_ready'] = True
    task['counts']['source_imports'] += 1
