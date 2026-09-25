"""Resolve different source values through the same persisted question channel."""
from .storage import digest, get, put

BLANK = '这些值暂不采用，留空'


def conflicting_groups(rows):
    groups = {}
    for row in rows:
        if not row.get('dropped'):
            groups.setdefault((row.get('entity_name'), row['key']), []).append(row)
    return [group for group in groups.values()
            if len({row['value'] for row in group}) > 1]


def apply_answer(group, answer):
    custom = answer.get('custom', '').strip()
    selected = answer.get('selected') or []
    choice = custom or (selected[0] if selected else '')
    if not choice:
        raise ValueError('冲突值尚未选择，未填写')
    values = sorted({row['value'] for row in group})
    options = {f'采用第 {i + 1} 项': value for i, value in enumerate(values)}
    if not custom and choice == BLANK:
        for row in group:
            row['dropped'] = True
        return
    if not custom and choice not in options:
        raise ValueError('冲突确认选项不属于当前材料，未采用')
    value = custom if custom else options[choice]
    kept = next((r for r in group if r['value'] == value), group[0])
    for row in group:
        row['dropped'] = row is not kept
    kept['value'] = value
    kept['resolved_by_user'] = True
    if custom:
        kept['quote'] += '；用户核实后填写：' + value


def prepare(store, task):
    task['source_conflict_groups'] = {}
    questions = []
    for group in conflicting_groups(task['rows']):
        first = group[0]
        evidence = sorted({(r['_source'], r['key'], r['value'], r.get('entity_name') or '',
                            r.get('quote') or '') for r in group})
        qid = 'source-conflict-' + digest([task['batch'], evidence])
        task['source_conflict_groups'][qid] = [r['n'] for r in group]
        reply = get(store.conn, 'office_v2_cache', qid)
        if reply is not None:
            apply_answer(group, reply)
            continue
        values = sorted({r['value'] for r in group})
        options = [{'label': BLANK, 'description': '尚未核实哪个正确时，建议先留空；不替你猜选其中一个值。'}]
        excerpts = []
        for i, value in enumerate(values):
            locations = [r for r in group if r['value'] == value]
            evidence_text = '；'.join(dict.fromkeys(
                f"文件《{r['_source']}》原文：{r['quote']}" for r in locations))
            options.append({'label': f'采用第 {i + 1} 项',
                            'description': f'{value}。{evidence_text}'})
            excerpts.append(f'第 {i + 1} 项：{value}。{evidence_text}')
        questions.append({'id': qid, 'header': '同一字段有不同值',
                          'question': f"主体：{first.get('entity_name') or '尚未确定'}；字段：{first['key']}。"
                          + ' '.join(excerpts)
                          + ' 请选择核实后的值，或自定义输入；不确定可以留空。',
                          'options': options})
    return questions


def save_answer(store, task, answer):
    numbers = task['source_conflict_groups'][answer['id']]
    group = [r for r in task['rows'] if r['n'] in numbers]
    apply_answer(group, answer)
    put(store.conn, 'office_v2_cache', answer['id'], answer)
