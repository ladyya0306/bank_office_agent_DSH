"""Bounded task views, including read-only original documents for recovery."""
import json
import re
from pathlib import Path
from . import mapping
from .storage import digest, inside

PAGE_CHARS = 5200
_XLSX_PLACEHOLDER_RE = re.compile(r'^[\s_＿—\-□☐]+$')


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
    return {key: (str(value)[:160] + '…' if key not in ('field', 'id', 'same_label_references') and len(str(value)) > 160 else value)
            for key, value in item.items() if key in ('field', 'label', 'id', 'value', 'entity', 'source', 'same_label_references')} | {
                'details_read': read, 'note': '本项较长，完整内容请按 details_read 读取；此处是摘要。'}


def _xlsx_same_label_references(task, template, slot, workbook_cache=None):
    """Show existing same-workbook label/value evidence; never infer a value."""
    target = slot.get('target') or {}
    if target.get('kind') != 'xlsx_cell' or not template:
        return []
    label = str(slot.get('label') or '').strip()
    work = task.get('last_work')
    if not label or not work:
        return []
    path = Path(work) / template
    if not path.is_file():
        return []
    cache_key = str(path.resolve())
    if workbook_cache is not None and cache_key in workbook_cache:
        workbook = workbook_cache[cache_key]
    else:
        try:
            from office_kit.target_validation import read_excel_template
            workbook = read_excel_template(path)
        except (OSError, KeyError, ValueError):
            workbook = None
        if workbook_cache is not None:
            workbook_cache[cache_key] = workbook
    if workbook is None:
        return []
    target_sheet = (workbook.worksheets[target['sheet']].title
                    if isinstance(target.get('sheet'), int) else target.get('sheet'))
    references = []
    for worksheet in workbook.worksheets:
        for row in worksheet.iter_rows():
            for cell in row:
                if worksheet.title == target_sheet and cell.coordinate == target.get('cell'):
                    continue
                value = cell.value
                if not isinstance(value, str) or value.startswith('='):
                    continue
                match = re.match(r'^\s*([^：:]+?)\s*[：:]\s*(.*?)\s*$', value, re.S)
                if not match or match.group(1).strip() != label:
                    continue
                text = match.group(2).strip()
                if not text or _XLSX_PLACEHOLDER_RE.fullmatch(text):
                    continue
                references.append({'sheet': worksheet.title, 'cell': cell.coordinate, 'text': value.strip()[:360]})
                if len(references) == 3:
                    return references
    return references


def _document_items(task, template, path=None):
    """Read one task-owned template on demand for mapping context."""
    path = path or (Path(task['last_work']) / template)
    if path.suffix.lower() == '.docx':
        from office_kit.target_validation import read_word_template
        engine = read_word_template(path)
        indexes = {}
        items = []
        for element, paragraph in engine._entries:
            part = engine._part_by_el.get(element,
                                          engine._part_by_el.get(paragraph._element, 'word/document.xml'))
            index = indexes.get(part, 0)
            indexes[part] = index + 1
            raw = engine.xml_for(paragraph)
            text = raw.text if raw is not None else paragraph.text
            for fragment in chunks(text, {'paragraph_index': index}):
                # chunks uses part for its numeric segment index; keep the
                # document's XML part address separately meaningful here.
                fragment['segment'] = fragment.pop('part')
                fragment['part'] = part
                items.append(fragment)
        return items
    if path.suffix.lower() == '.xlsx':
        from office_kit.target_validation import read_excel_template
        workbook = read_excel_template(path)
        items = []
        for worksheet in workbook.worksheets:
            for row in worksheet.iter_rows():
                for cell in row:
                    if cell.value not in (None, ''):
                        items.extend(chunks(cell.value, {'sheet': worksheet.title, 'cell': cell.coordinate}))
        return items
    raise ValueError('当前模板不是可读取的 Word 或 Excel 文件')


def page(task, query):
    section = query.get('section', 'positions')
    offset = query.get('offset', 0)
    if not isinstance(offset, int) or isinstance(offset, bool) or offset < 0:
        raise ValueError('分页 offset 必须为非负整数')
    version = revision(task)
    if section == 'source_document':
        from office_kit.store_v2 import sha256_file
        name = query.get('source')
        if name not in task.get('source', []):
            raise ValueError('只能读取本任务登记的来源')
        version = digest([version, name, sha256_file(inside(Path(task['last_work']), name))])
    if query.get('revision') and query['revision'] != version:
        raise ValueError('来源或位置已经变化，请重新读取第一页；未修改位置')
    names = [n for n, r in task.get('documents', {}).items() if r.get('status') == 'needs_mapping']
    template = query.get('template') or (names[0] if names else None)
    record = task.get('documents', {}).get(template)
    if section == 'fields':
        items = [shorten(f, {'section': 'source_details', 'field': f['field']}) for f in facts(task)]
    elif section == 'source_document':
        from .source_evidence import document_items
        items = []
        for original in document_items(Path(task['last_work']), task, query.get('source')):
            position = {key: value for key, value in original.items() if key != 'text'}
            for fragment in chunks(original['text'], {}):
                fragment['segment'] = fragment.pop('part')
                items.append({**position, **fragment})
    elif section == 'source':
        seen, items = set(), []
        for row in task.get('rows', []):
            key = (row.get('_source'), row.get('line'), row.get('quote'))
            if key in seen:
                continue
            seen.add(key)
            item = {'source': row.get('_source'), 'line': row.get('line'),
                    'original': row.get('quote'), 'owner': row.get('entity_name'),
                    'context': row.get('context')}
            items.extend([item] if size(item) <= 2100 else chunks(item, {'source_row': len(seen)}))
    elif section == 'source_details':
        match = next((f for f in facts(task) if query.get('field') in [f['field'], *f['aliases']]), None)
        if match is None:
            raise ValueError('找不到该来源字段')
        items = list(chunks(match, {'field': match['field']}))
    elif section == 'templates':
        items = []
        for name in task.get('targets', list(task.get('documents', {}))):
            record = task.get('documents', {}).get(name)
            items.append({'template': name, 'status': record.get('status') if record else 'not_parsed',
                          'positions': len(record.get('unmapped_slots', [])) if record else None,
                          'current_positions': len(record.get('current_positions', [])) if record else None})
    elif section == 'issues':
        items = list(chunks(task.get('issues', []), {}))
    elif section in ('positions', 'current_positions', 'context', 'document', 'output'):
        if section in ('document', 'output') and not template:
            raise ValueError('读取完整模板需要指定本任务的 template')
        if record is None and template is not None:
            if template not in task.get('targets', []):
                raise ValueError('只能读取本任务的模板；请使用 templates 页列出的完整路径')
            if section != 'document':
                raise ValueError('该模板属于本任务，但尚未建立填写位置；请先处理来源解析问题。正文可用 document 页读取。')
        record = record or {}
        if section == 'positions':
            workbook_cache = {}
            items = []
            for slot in record.get('unmapped_slots', []):
                item = mapping.visible_slot(slot)
                references = _xlsx_same_label_references(task, template, slot, workbook_cache)
                if references:
                    item['same_label_references'] = references
                items.append(shorten(item, {'section': 'context', 'template': template, 'slot_id': slot['id']}))
        elif section == 'current_positions':
            items = list(chunks(record.get('current_positions', []), {'template': template}))
        elif section == 'context':
            slot = next((s for s in record.get('slots', []) if s['id'] == query.get('slot_id')), None)
            if slot is None:
                raise ValueError('找不到该填写位置')
            detail = {'label': slot.get('label'), 'context': slot.get('context'),
                      'original': slot.get('target', {}).get('expected_text')}
            items = list(chunks(detail, {'slot_id': slot['id']}))
            references = _xlsx_same_label_references(task, template, slot)
            if references:
                items.append({'slot_id': slot['id'], 'same_label_references': references})
        elif section == 'document':
            items = _document_items(task, template)
        else:
            output_path = None
            if record.get('status') != 'completed':
                items = [{'template': template, 'reason': '该模板当前没有已完成产物'}]
            elif not record.get('output'):
                items = [{'template': template, 'reason': '该模板没有当前产物路径'}]
            else:
                try:
                    output_path = inside(Path(task['last_work']), record['output'])
                except (OSError, ValueError):
                    items = [{'template': template, 'reason': '当前产物不存在或不在本工作区内'}]
                else:
                    items = _document_items(task, template, output_path)
    else:
        raise ValueError('不支持的 mapping_read.section')
    selected = []
    for item in items[offset:]:
        if selected and size(selected + [item]) > PAGE_CHARS:
            break
        selected.append(item)
    next_offset = offset + len(selected)
    base = {**query, 'section': section, 'revision': version}
    if section in ('positions', 'current_positions', 'context', 'document', 'output'):
        base['template'] = template
    result = {'section': section, 'template': template if section in ('positions', 'current_positions', 'context', 'document', 'output') else None,
            'offset': offset, 'total': len(items), 'items': selected, 'revision': version,
            'next': {**base, 'offset': next_offset} if next_offset < len(items) else None}
    if section == 'source_document':
        from office_kit.store_v2 import sha256_file
        result['source'] = query['source']
        result['source_sha256'] = sha256_file(inside(Path(task['last_work']), query['source']))
    if section == 'output' and 'output_path' in locals() and output_path is not None:
        result['output'] = str(output_path)
    if section == 'positions':
        result.update(scope='unmapped_only',
                      already_mapped_fields=list(dict.fromkeys(
                          item['field'] for item in (record or {}).get('current_positions', [])
                          if item.get('field')))[:20],
                      note='items 只列尚未关联的空位，不含已自动关联或明确留空的位置。数量少不表示漏识别；已关联位置可用 current_positions 查看，不必重新解压模板。')
    return result


def initial(task):
    fields = page(task, {'section': 'fields'})
    positions = page(task, {'section': 'positions'})
    templates = page(task, {'section': 'templates'})
    return {'available_fields': fields.pop('items'),
            'mapping_requests': ([{'template': positions['template'], 'positions': positions.pop('items')}]
                                 if positions['template'] else []),
            'mapping_pages': {'fields': fields, 'positions': positions, 'templates': templates},
            'mapping_overview': {'templates': len(task.get('documents', {})),
                                 'unmapped_positions': sum(len(r.get('unmapped_slots', [])) for r in task.get('documents', {}).values()),
                                 'source_rows': len(task.get('rows', []))},
            'source_read': {'section': 'source'},
            'mapping_help': 'positions 只含尚未关联的空位，不包括已自动关联或明确留空的位置；已关联位置用 current_positions 查看。本页只含第一份模板的部分待关联位置，不是全部模板统计。来源别名 aliases 均可作为 field。源材料原文及归属上下文用 mapping_read:{"section":"source"}；字段续页和位置续页原样传 next。其他模板用 {"section":"positions","template":"文件名"}。位置完整原文用 context+template+slot_id。以上均读取已缓存证据，不必用命令解压文档或再解析来源。'}
