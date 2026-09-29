"""User-confirmed extraction results and reusable script/position records.

Scripts are retained, not exec'd by the office process. DSH's existing tools
remain the execution boundary. Facts are reusable only for identical bytes;
changing data requires a fresh, evidenced extraction, never cached answers.
"""
from pathlib import Path
from office_kit.store_v2 import sha256_file
from office_kit.rule_pack import build_pack, load_pack
from .storage import digest, get, put, inside
from .source_evidence import validate_updates

PREFIX = 'method:'


def override_key(name, fingerprint):
    return 'confirmed-source:' + digest([name, fingerprint])


def summary(item):
    return {k: item[k] for k in ('id', 'name', 'script', 'source_count', 'template_count') if k in item}


def read(store, work, query):
    if query.get('action') == 'list':
        import json
        return {'methods': [summary(json.loads(r[0])) for r in store.conn.execute(
            'SELECT payload FROM office_v2_cache WHERE id LIKE ? ORDER BY id', (PREFIX + '%',))]}
    item = get(store.conn, 'office_v2_cache', PREFIX + str(query.get('id', '')))
    if item is None:
        raise ValueError('找不到本工作区已确认的方法；请先列出 list')
    result = summary(item)
    result['source_rules'] = item['source_updates']
    result['target_rules'] = item.get('target_pack')
    script = item.get('script')
    if script:
        path = inside(work, script['path'])
        if sha256_file(path) != script['sha256']:
            raise ValueError('保存的脚本内容已被修改；请重新提交并确认此版本，未执行脚本')
        result['script_path'] = str(path)
    result['next_action'] = ('相同来源可 apply 复用已确认结果；来源改变时，读取已保存脚本并通过 DSH 现有执行工具重新运行，'
                             '把新结果作为 source_updates 提交。不要修改共享程序，也不要直接写事实数据库。')
    return result


def snapshot(store, work, task, proposal):
    name = proposal.get('name')
    if not isinstance(name, str) or not name.strip() or len(name) > 120:
        raise ValueError('方法需要 1 至 120 字的名称')
    updates = proposal.get('source_updates') or []
    rows = validate_updates(work, task, updates) if updates else []
    updates = [{**update, 'source_sha256': update['source_sha256'].lower()} for update in updates]
    paths = [inside(work, n) for n in task['targets'] if store.rules_for(inside(work, n))]
    pack = build_pack(store, paths) if proposal.get('save_target_rules') and paths else None
    if proposal.get('save_target_rules') and not pack:
        raise ValueError('当前没有可保存的填写规则；先完成位置关联，再保存')
    item = {'name': name.strip(), 'source_updates': updates, 'rows': rows,
            'target_pack': pack, 'source_count': len(updates),
            'template_count': len(pack['templates']) if pack else 0}
    if proposal.get('script_path'):
        path = inside(work, proposal['script_path'])
        if path.suffix.lower() not in ('.py', '.js', '.mjs', '.ps1'):
            raise ValueError('可保存 Python、JavaScript 或 PowerShell 文本脚本')
        if path.stat().st_size > 128 * 1024:
            raise ValueError('脚本超过 128 KiB，请保留本次识别逻辑后再提交')
        content = path.read_text(encoding='utf-8-sig')
        item['script'] = {'path': path.relative_to(work).as_posix(), 'sha256': sha256_file(path),
                          'suffix': path.suffix.lower()}
        item['script_content'] = content
    if not rows and not pack and 'script' not in item:
        raise ValueError('请提供来源识别结果、脚本或已有填写规则中的至少一项')
    item['id'] = digest(item)
    return item


def propose(store, work, task, proposal):
    item = snapshot(store, work, task, proposal)
    old = get(store.conn, 'office_v2_cache', PREFIX + item['id'])
    if old:
        apply_method(store, work, task, {'id': item['id']})
        return
    task['pending_method'] = {'proposal': proposal, 'snapshot': item}
    task['learning'] = {'status': 'awaiting_confirmation', **summary(item)}
    questions(task)


def questions(task):
    item = task['pending_method']['snapshot']
    detail = []
    if item.get('script'):
        detail.append('保存脚本：' + item['script']['path'] + '；版本 ' + item['script']['sha256'][:12]
                      + '。保存不会直接运行它；再次执行仍使用现有执行工具。')
    for update in item['source_updates']:
        detail.append('替换来源识别结果：' + update['source'])
        for row in [r for r in item['rows'] if r['_source'] == update['source']]:
            detail.append(f"{row.get('entity_name') or '不绑定主体'} / {row['key']} = {row['value']}"
                          + (f"；角色：{row['role']}" if row.get('role') else '')
                          + f"；依据 {row['line']}：{row['quote']}"
                          + (f"；转换/计算（请核对）：{row['derivation']}" if row.get('derivation') else ''))
    for template in (item.get('target_pack') or {}).get('templates', []):
        detail.append('保存模板位置：' + template['name'])
        detail.extend(f"{r['field']} → {r['target']}" for r in template['rules'])
    # Separate bounded cards keep all proposed fields visible; one answer cannot
    # silently approve a tail omitted from a truncated summary.
    pages, current = [], []
    for line in detail:
        if current and len('\n'.join(current + [line])) > 3500:
            pages.append(current); current = []
        current.append(line)
    pages.append(current)
    task['status'] = 'awaiting_method'
    task['questions'] = [{'id': f"method-{item['id']}-{index}", 'header': '保存识别方法',
                          'question': f"方法：{item['name']}（{index + 1}/{len(pages)}）。确认以下内容供本工作区复用？\n" + '\n'.join(page),
                          'options': [{'label': '确认保存', 'description': '保存本页；全部页确认后才生效'},
                                      {'label': '暂不保存', 'description': '不收录本次方法和结果'}],
                          'multiSelect': False} for index, page in enumerate(pages)]


def confirm(store, work, task, answers):
    pending = task['pending_method']
    if any(a.get('selected') != ['确认保存'] or a.get('custom') for a in answers):
        task.pop('pending_method')
        task['learning'] = {
            'status': 'not_saved',
            'next_action': '本次方法未保存，原任务和已有资料均保留。要修订方法请使用同一 task_id 再次提出方法；'
                           '要继续原任务请使用同一 task_id 恢复，当前不会重复弹出旧来源问题。',
        }
        # This is a cancellation of the save-method request, not an answer to
        # the source questions that were visible before proposing it.  Keep the
        # task and its evidence intact, but make this resume call terminal so
        # runner does not immediately reopen those old questions.
        task.update(status='cancelled', questions=[])
        return
    try:
        item = snapshot(store, work, task, pending['proposal'])
    except (ValueError, OSError):
        task.pop('pending_method')
        task['learning'] = {'status': 'needs_refresh',
                            'next_action': '确认依据已失效；请读取当前文件，重新 propose，不要重答旧问题。'}
        task.update(status='needs_source_update', questions=[])
        raise
    if item['id'] != pending['snapshot']['id']:
        task.pop('pending_method')
        task['learning'] = {'status': 'needs_refresh',
                            'next_action': '文件或规则已改变，请重新 propose 当前方法，不要重答旧问题。'}
        task.update(status='needs_source_update', questions=[])
        raise ValueError('确认期间来源、脚本或填写规则已改变；未保存，请重新提交当前版本')
    if item.get('script'):
        # Content-addressed, workspace-local copy; never replace a user's file.
        folder = work / '.office-methods' / item['id']
        if not folder.resolve().is_relative_to(work):
            raise ValueError('方法保存目录实际位置超出工作区')
        folder.mkdir(parents=True, exist_ok=True)
        path = folder / ('script' + item['script']['suffix'])
        if not path.resolve().is_relative_to(work):
            raise ValueError('脚本保存位置超出工作区')
        raw = inside(work, item['script']['path']).read_bytes()
        import hashlib
        if hashlib.sha256(raw).hexdigest() != item['script']['sha256']:
            raise ValueError('脚本在确认后发生改变；未保存此版本')
        if path.exists():
            if path.read_bytes() != raw:
                raise ValueError('方法保存目录中存在不同脚本；未覆盖')
        else:
            with path.open('xb') as stream:
                stream.write(raw)
        item['script']['path'] = path.relative_to(work).as_posix()
    item.pop('script_content', None)
    # The audit event is required before making this record reusable.
    store.event('method_confirmed', batch_no=task['batch'], target=task['id'],
                payload={'method_id': item['id'], 'name': item['name'], 'answers': answers}, actor='user')
    put(store.conn, 'office_v2_cache', PREFIX + item['id'], item)
    task.pop('pending_method')
    apply_method(store, work, task, {'id': item['id']})


def apply_method(store, work, task, query):
    item = get(store.conn, 'office_v2_cache', PREFIX + str(query.get('id', '')))
    if item is None:
        raise ValueError('方法尚未由使用者确认保存')
    read(store, work, {'id': item['id']})  # Verify archived script, if present.
    updates, changed = [], []
    for update in item['source_updates']:
        name = update['source']
        if name not in task['source']:
            matches = [n for n in task['source']
                       if sha256_file(inside(work, n)) == update['source_sha256']]
            if len(matches) == 1:
                name = matches[0]
        if name not in task['source'] or sha256_file(inside(work, name)) != update['source_sha256']:
            changed.append({'source': name, 'read': {'section': 'source_document', 'source': name}
                            if name in task['source'] else None})
        updates.append({**update, 'source': name})
    if changed:
        task['learning'] = {'status': 'needs_refresh', **summary(item), 'changed_sources': changed,
                            'next_action': '来源已变化，旧值不适用。使用 read 返回的已保存脚本重新读取当前文件，'
                                           '再 propose 新的 source_updates；不用修改共享工具或重新保存未改变的脚本。'}
        task.update(status='needs_source_update', source_ready=False, questions=[], available_fields=[])
        return False
    rows = validate_updates(work, task, updates) if updates else []
    pack = item.get('target_pack')
    selected = []
    if pack:
        for entry in pack['templates']:
            matches = [inside(work, n) for n in task['targets']
                       if sha256_file(inside(work, n)) == entry['sha256']]
            if len(matches) != 1:
                raise ValueError('模板格式已改变或有多个相同模板，请重新确认位置：' + entry['name'])
            selected.append((matches[0], {**entry, 'name': matches[0].name}))
    # All evidence/template hashes are checked before any reusable result writes.
    for update in updates:
        source_rows = [r for r in rows if r['_source'] == update['source']]
        put(store.conn, 'office_v2_cache', override_key(update['source'], update['source_sha256']),
            {'rows': source_rows, 'issues': [], 'method_id': item['id']})
    conflicts = 0
    for path, entry in selected:
        result = load_pack(store, {**pack, 'templates': [entry]}, [path], task['batch'])
        conflicts += result['conflicts_skipped']
    task['learning'] = {'status': 'applied', **summary(item)}
    if conflicts:
        task['learning'].update(conflicts_skipped=conflicts,
                                note='当前已有不同位置规则，已保留；请核对 current_positions，不重复提交保存版本')
    store.event('method_reused', batch_no=task['batch'], target=task['id'],
                payload={'method_id': item['id']}, actor='workflow')
    return True
