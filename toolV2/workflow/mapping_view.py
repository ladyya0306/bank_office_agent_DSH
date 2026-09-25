"""Small, repeatable pages from the existing task; never parse documents here."""
import json
from . import mapping
from .storage import digest

PAGE_CHARS = 5200


def size(value):
    return len(json.dumps(value, ensure_ascii=False, separators=(',', ':')))


def revision(task):
    return digest([task.get('source_signature'), task.get('available_fields'),
                   [[name, r.get('slot_hash'), r.get('slot_version'),
                     r.get('current_positions'), r.get('blank_slots')]
                    for name, r in task.get('documents', {}).items()]])


def facts(task):
    """Aliases retain their selectable keys, while evidence appears only once."""
    grouped = {}
    for row in task.get('available_fields', []):
        candidates = row.get('candidates') or [row]
        for candidate in candidates:
            item = {'value': candidate.get('value'),
                    'entity': candidate.get('entity_name', candidate.get('entity', row.get('entity'))),
                    'source': candidate.get('provenance', candidate.get('source', row.get('source')))}
            fact_id = candidate.get('fact_id', candidate.get('id'))
            key = digest([fact_id, candidate.get('entity_id'), item['value'], item['source']]
                         if fact_id is not None else [item, candidate.get('entity_id')])
            if key not in grouped:
                grouped[key] = {'field': row['field'], 'aliases': [], **item}
            elif row['field'] != grouped[key]['field'] and row['field'] not in grouped[key]['aliases']:
                grouped[key]['aliases'].append(row['field'])
            if not grouped[key].get('entity') and item.get('entity'):
                grouped[key]['entity'] = item['entity']
    return list(grouped.values())


def chunks(value, context):
    """Unusually long evidence is still readable text, not serialized JSON to decode."""
    if isinstance(value, dict):
        for key, child in value.items():
            yield from chunks(child, {**context, 'property': context.get('property', '') + '.' + str(key)})
    elif isinstance(value, list):
        for index, child in enumerate(value):
            yield from chunks(child, {**context, 'property': context.get('property', '') + f'[{index}]'})
    else:
        text = '' if value is None else str(value)
        parts = [text[i:i+900] for i in range(0, len(text), 900)] or ['']
        for index, part in enumerate(parts):
            yield {**context, 'text': part, 'part': index + 1, 'parts': len(parts)}


def shorten(item, read):
    if size(item) <= 2100:
        return item
    # The original is always accessible through the same tool, with an explicit pointer.
    return {key: (str(value)[:160] + '…' if key not in ('field', 'id') and len(str(value)) > 160 else value)
            for key, value in item.items() if key in ('field', 'label', 'id', 'value', 'entity', 'source')} | {
                'details_read': read, 'note': '本项较长，完整内容请按 details_read 读取；此处是摘要。'}


def page(task, query):
    section = query.get('section', 'positions')
    offset = query.get('offset', 0)
    if not isinstance(offset, int) or isinstance(offset, bool) or offset < 0:
        raise ValueError('分页 offset 必须为非负整数')
    version = revision(task)
    if query.get('revision') and query['revision'] != version:
        raise ValueError('来源或位置已经变化，请重新读取第一页；未修改位置')
    names = [n for n, r in task.get('documents', {}).items() if r.get('status') == 'needs_mapping']
    template = query.get('template') or (names[0] if names else None)
    record = task.get('documents', {}).get(template)
    if section == 'fields':
        items = [shorten(f, {'section': 'source_details', 'field': f['field']}) for f in facts(task)]
    elif section == 'source_details':
        match = next((f for f in facts(task) if query.get('field') in [f['field'], *f['aliases']]), None)
        if match is None:
            raise ValueError('找不到该来源字段')
        items = list(chunks(match, {'field': match['field']}))
    elif section == 'templates':
        items = [{'template': n, 'status': r.get('status'), 'positions': len(r.get('unmapped_slots', [])),
                  'current_positions': len(r.get('current_positions', []))}
                 for n, r in task.get('documents', {}).items()]
    elif section == 'issues':
        items = list(chunks(task.get('issues', []), {}))
    elif section in ('positions', 'current_positions', 'context'):
        if record is None and template is not None:
            raise ValueError('只能读取本任务的模板')
        record = record or {}
        if section == 'positions':
            items = [shorten(mapping.visible_slot(s), {'section': 'context', 'template': template, 'slot_id': s['id']})
                     for s in record.get('unmapped_slots', [])]
        elif section == 'current_positions':
            items = list(chunks(record.get('current_positions', []), {'template': template}))
        else:
            slot = next((s for s in record.get('slots', []) if s['id'] == query.get('slot_id')), None)
            if slot is None:
                raise ValueError('找不到该填写位置')
            items = list(chunks({'label': slot.get('label'), 'context': slot.get('context'),
                                 'original': slot.get('target', {}).get('expected_text')}, {'slot_id': slot['id']}))
    else:
        raise ValueError('不支持的 mapping_read.section')
    selected = []
    for item in items[offset:]:
        if selected and size(selected + [item]) > PAGE_CHARS:
            break
        selected.append(item)
    next_offset = offset + len(selected)
    base = {**query, 'section': section, 'revision': version}
    if section in ('positions', 'current_positions', 'context'):
        base['template'] = template
    return {'section': section, 'template': template if section in ('positions', 'current_positions', 'context') else None,
            'offset': offset, 'total': len(items), 'items': selected, 'revision': version,
            'next': {**base, 'offset': next_offset} if next_offset < len(items) else None}


def initial(task):
    fields = page(task, {'section': 'fields'})
    positions = page(task, {'section': 'positions'})
    templates = page(task, {'section': 'templates'})
    return {'available_fields': fields.pop('items'),
            'mapping_requests': ([{'template': positions['template'], 'positions': positions.pop('items')}]
                                 if positions['template'] else []),
            'mapping_pages': {'fields': fields, 'positions': positions, 'templates': templates},
            'mapping_help': '来源别名 aliases 均可作为 field。只读第一页所列模板；其余模板用 mapping_read={section:positions,template:文件名}，续页原样传 next。完整原文用 section:context+template+slot_id；已有规则用 current_positions。无需读临时文件或再次解析源文件。'}
