"""Reuse saved fill answers and group identical questions across documents."""
from office_kit.fill_decisions import fingerprint, saved_choices, save_choices
from .storage import digest

ACCEPT = '采用建议值（推荐）'
BLANK = '留空'


def location(place):
    target = place['target']
    if target.get('expected_text') is not None:
        text = target['expected_text']
        start, end = int(target['span_start']), int(target['span_end'])
        label = f"原文：{text[max(0,start-45):start]}【填写处】{text[end:end+60]}"
        if target.get('sheet'):
            label = f"{target['sheet']} 的 {target['cell']}，{label}"
    elif target.get('cell'):
        label = f"{target.get('sheet', '工作表')} 的 {target['cell']} 单元格"
    elif target.get('anchor'):
        label = f"“{target['anchor']}”对应的填写处"
        if str(target.get('part', '')).startswith('word/header'):
            label = '页眉内' + label
        elif str(target.get('part', '')).startswith('word/footer'):
            label = '页脚内' + label
    elif target.get('kind') == 'cell':
        label = f"第 {int(target['table']) + 1} 个表格，第 {int(target['row']) + 1} 行第 {int(target['col']) + 1} 格，{place.get('label', '')}"
    else:
        label = place.get('label', '指定填写处')
    return f"{place['template']}（{label}）"


def questions(task):
    groups = {}
    for template, record in task['documents'].items():
        if record.get('status') in ('completed', 'needs_mapping') or not record.get('plan'):
            continue
        plan = record['plan']
        _, missing = saved_choices(plan)
        for row in plan['rows']:
            if row.get('n') not in missing:
                continue
            # Unresolved owners cannot share an answer merely because their
            # candidate lists happen to be identical.
            owner = row.get('subject_eid') or row.get('entity_id')
            scope = [row.get('subject_scope'), owner if owner is not None else template]
            key = digest([scope, row['field'], row.get('value'), row.get('entity_name'),
                          row.get('ambiguous'), row.get('candidates'), row.get('provenance')])
            group = groups.setdefault(key, {'row': row, 'places': []})
            group['places'].append({'template': template, 'n': row['n'],
                                    'fingerprint': fingerprint(plan, row), 'target': row['target'],
                                    'label': row.get('label', row['field'])})
    output = []
    task['question_groups'] = {}
    for group in groups.values():
        row = group['row']
        qid = 'fill-' + digest([task['batch'], group['places']])
        if row.get('ambiguous'):
            options = [{'label': f"使用主体 #{c['entity_id']} {c['entity_name']}",
                        'description': f"值：{c['value']}"} for c in row.get('candidates', [])]
        elif row.get('value') not in (None, ''):
            options = [{'label': ACCEPT, 'description': f"填写：{row['value']}"}]
        else:
            options = []
        options.append({'label': BLANK, 'description': '本卡列出的位置全部留空。'})
        places = '；'.join(location(p) for p in group['places'])
        if row.get('ambiguous'):
            suggestion = '源材料中有多个主体的值，请选择本表要使用的那一项，并非资料缺失'
        elif row.get('value') not in (None, ''):
            suggestion = f"建议填写：{row['value']}（来源主体：{row.get('entity_name') or '业务公共信息'}）"
        else:
            suggestion = f"未找到{row.get('subject_label') or '本表对应主体'}的这项信息，可留空"
        output.append({'id': qid, 'header': row['field'],
                       'question': f"填写位置：{places}。{suggestion}。自定义回答请直接填写要放入这些位置的文字；不同位置需不同内容时，请勿合写为一句说明。",
                       'options': options})
        task['question_groups'][qid] = group
    return output


def validate(questions, answers):
    if not isinstance(answers, list) or not answers:
        return False
    expected = {q['id']: q for q in questions}
    if len({a.get('id') for a in answers}) != len(answers):
        raise ValueError('同一个问题被重复提交，未保存')
    if set(expected) != {a.get('id') for a in answers}:
        raise ValueError('回答缺失或不属于当前问题，未保存')
    for a in answers:
        selected, custom = a.get('selected', []), a.get('custom', '')
        if not isinstance(custom, str) or not isinstance(selected, list):
            raise ValueError('回答格式无效')
        if not custom.strip() and (len(selected) != 1 or selected[0] not in
                                   {o['label'] for o in expected[a['id']]['options']}):
            raise ValueError('请选择一个选项或填写答案')
    return True


def save(task, answers):
    choices = {t: saved_choices(r['plan'])[0] for t, r in task['documents'].items()
               if r.get('plan') and r.get('status') not in ('completed', 'needs_mapping')}
    for answer in answers:
        group = task['question_groups'][answer['id']]
        custom = answer.get('custom', '').strip()
        pick = custom or answer['selected'][0]
        for place in group['places']:
            c, n = choices[place['template']], str(place['n'])
            if pick in (BLANK, '暂不填写', '先不填', '不知道', '不确定'):
                c['blank'].append(n)
            elif custom:
                c['new'].append(f'{n}={custom}')
            elif group['row'].get('ambiguous'):
                candidate = next((x for x in group['row'].get('candidates', []) if pick in
                    (x['entity_name'], str(x['entity_id']), f"#{x['entity_id']}",
                     f"使用主体 #{x['entity_id']} {x['entity_name']}")), None)
                if candidate is None:
                    raise ValueError('请明确选择提供该值的主体，未保存')
                c['use'].append(f"{n}=#{candidate['entity_id']}")
            elif pick == ACCEPT:
                c['select'] = ','.join(filter(None, [c['select'], n]))
            else:
                raise ValueError('未知填写选项')
    for template, choice in choices.items():
        save_choices(task['documents'][template]['plan'], choice)


def execution_signature(plan):
    selections, missing = saved_choices(plan)
    rows = [{k: r.get(k) for k in ('template_sha256', 'field', 'label', 'target', 'value',
                                   'entity_name', 'subject_eid', 'subject_scope', 'provenance', 'decision', 'candidates', 'source_kind')}
            for r in plan['rows']]
    return digest([plan['batch_no'], rows, selections, missing])
