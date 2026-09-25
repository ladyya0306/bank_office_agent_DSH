"""Deterministic fill-task entry. Model calls one tool; code advances it."""
from __future__ import annotations
import json
import os
from pathlib import Path
import subprocess
import sys
import uuid
from office_kit.store_v2 import StoreV2, allocate_batch, sha256_file, valid_batch
from office_kit.workroot import init_workroot, is_workroot
from office_kit.cli import build_parser, _dispatch
from office_kit.harness import build_fill_plan, current_docx_target_errors, current_xlsx_gaps
from . import storage as state, source, review, mapping

ROOT = Path(__file__).resolve().parents[1]
WorkflowError = ValueError
POSITION_PARSE_VERSION = 4
SUBJECT_PLAN_VERSION = 3


def command(argv):
    return json.loads(_dispatch(build_parser().parse_args(argv)).to_json())


def save(store, task):
    state.put(store.conn, 'office_v2_task', task['id'], task)


def response(work, task):
    results = []
    for name, record in task.get('documents', {}).items():
        result = {k: record[k] for k in ('status', 'error', 'filled', 'failed', 'coverage', 'blank_slots') if k in record}
        result['template'] = name
        if record.get('output') and record.get('status') == 'completed':
            result['output'] = str(work / record['output'])
        results.append(result)
    return {'ok': task['status'] != 'failed', 'status': task['status'], 'task_id': task['id'],
            'work': str(work), 'batch': task['batch'], 'questions': task.get('questions', []),
            'issues': task.get('issues', []), 'results': results, 'counters': task['counts'],
            'report': str(work / task['report']) if task.get('report') else None,
            'mapping_requests': [{'template': name, 'positions': [mapping.visible_slot(s) for s in record.get('unmapped_slots', [])],
                                  'current_positions': record.get('current_positions', [])}
                                 for name, record in task['documents'].items() if record.get('status') == 'needs_mapping'],
            'available_fields': task.get('available_fields', []) if any(
                r.get('status') == 'needs_mapping' for r in task['documents'].values()) else [],
            'mapping_revision': state.digest([[name, r.get('current_positions'), r.get('blank_slots')]
                                              for name, r in task['documents'].items()]),
            'unsigned': True, 'note': ('部分源材料尚未解析，请查看 issues 中的文件和原文；已生成文件仅包含已处理内容。'
                                     if task.get('source_issues') else '')
                                    + '填报完成的文件供用户复核；未代用户签核。'}


def rebase(store, work, task):
    old = task.get('last_work')
    if not old or old == str(work):
        return
    for table, column in [('template', 'path'), ('source', 'path'), ('source', 'copy_path')]:
        for row in list(store.conn.execute(f'SELECT id,{column} FROM {table}')):
            try:
                relative = Path(row[1]).relative_to(old)
            except (ValueError, TypeError):
                continue
            new = work / relative
            if new.exists():
                store.conn.execute(f'UPDATE {table} SET {column}=? WHERE id=?', (str(new), row[0]))
    store.conn.commit()
    task['last_work'] = str(work)


def start(store, work, data):
    lists = []
    for key in ('source', 'targets'):
        names = data.get(key)
        if not isinstance(names, list) or not names or not all(isinstance(p, str) for p in names):
            raise ValueError(f'{key} 必须是非空文件路径数组')
        lists.append(sorted({state.inside(work, p).relative_to(work).as_posix() for p in names}))
    sources, targets = lists
    if set(sources) & set(targets):
        raise ValueError('源文件与目标模板不能是同一个文件')
    if any(Path(p).suffix.lower() not in ('.docx', '.xlsx') for p in targets):
        raise ValueError('当前填报入口支持 Word .docx 与 Excel .xlsx 目标模板')
    # An unrelated diagnostic task must not switch the batch of an existing
    # multi-file job when a new conversation asks to continue the same files.
    matching = None
    if not data.get('batch'):
        for saved in store.conn.execute('SELECT payload FROM office_v2_task ORDER BY rowid DESC'):
            previous = json.loads(saved[0])
            if previous.get('source') == sources and previous.get('targets') == targets:
                matching = previous
                break
    batch = data.get('batch') or (matching and matching['batch']) or store.current_batch() or allocate_batch(store.conn)
    if not valid_batch(batch):
        raise ValueError('批次格式应为 YYYYMMDD-NN；继续同一业务可不填')
    task_id = state.digest([batch, sources, targets])[:24]
    task = state.get(store.conn, 'office_v2_task', task_id)
    if task is None:
        task = {'id': task_id, 'batch': batch, 'source': sources, 'targets': targets,
                'documents': {}, 'counts': {'source_reads': 0, 'source_imports': 0,
                'position_proposals': 0, 'previews': 0, 'fill_processes': 0}, 'last_work': str(work)}
    rebase(store, work, task)
    return task


def prepare_documents(store, work, task):
    task['issues'] = list(task.get('source_issues', []))
    facts = [dict(r) for r in store.conn.execute('SELECT id,key,value,entity_id,provenance,source_kind,status FROM fact WHERE batch_no=? AND superseded_by IS NULL ORDER BY id', (task['batch'],))]
    roles = [tuple(r) for r in store.conn.execute('SELECT * FROM case_role ORDER BY id')]
    store.batch_scope = task['batch']
    catalog = mapping.fact_catalog(store)
    task['available_fields'] = mapping.visible_facts(catalog)
    for name in task['targets']:
        path = state.inside(work, name)
        prior = task['documents'].get(name, {})
        record = dict(prior)
        try:
            store.batch_scope = task['batch']
            store.register_template(path, task['batch'])
            if (record.get('blank_source_signature') != task.get('source_signature')
                    or record.get('blank_template_hash') != sha256_file(path)):
                record['blank_slots'] = {}
                record['blank_source_signature'] = task.get('source_signature')
                record['blank_template_hash'] = sha256_file(path)
            if (prior.get('proposed_version') != POSITION_PARSE_VERSION
                    or prior.get('proposed_hash') != sha256_file(path)):
                command(['db-propose', str(path), '--work', str(work), '--batch', task['batch']])
                record['proposed_hash'] = sha256_file(path)
                record['proposed_version'] = POSITION_PARSE_VERSION
                task['counts']['position_proposals'] += 1
            mapping.auto_map(store, path, record, catalog, task['batch'])
            uncovered = mapping.inspect(store, path, record, catalog)
            record['current_positions'] = [{'field': r['field'], 'target': r['target']} for r in store.rules_for(path)]
            inputs = state.digest([POSITION_PARSE_VERSION, SUBJECT_PLAN_VERSION, sha256_file(path), store.rules_for(path), facts, roles, record.get('blank_slots')])
            if prior.get('inputs') == inputs and prior.get('plan'):
                plan = prior['plan']
            else:
                plan = build_fill_plan(store, [path], batch_no=task['batch'],
                                       run_id=prior.get('plan', {}).get('run_id') or f"{task['batch']}-V2-{uuid.uuid4().hex[:12]}")
                task['counts']['previews'] += 1
            plan.update(db_path=str(work / 'db/workflow.db'), template_files=[str(path)])
            plan['position_inventory'] = {'template_hash': sha256_file(path),
                                          'blank_targets': [s['target'] for s in record.get('slots', [])
                                                            if s['id'] in record.get('blank_slots', {})]}
            if prior.get('inputs') == inputs:
                errors = prior.get('validation_errors', [])
            else:
                errors = [r for r in plan['rows'] if r['kind'] == 'template_note']
                errors += current_docx_target_errors([path], plan)
                errors += current_xlsx_gaps(store, [path], plan)
                for rule in store.rules_for(path):
                    try:
                        mapping.validate_target(path, rule['target'])
                    except Exception as exc:
                        errors.append({'field': rule['field'], 'reason': str(exc),
                                       'target': rule['target']})
            if record.get('slots') and not uncovered and not store.rules_for(path):
                # User requested missing/uncertain data remain blank. A template
                # whose positions were all accounted for can be copied unchanged.
                plan['rows'] = []
                errors = []
            record.update(inputs=inputs, validation_errors=list(errors))
            if uncovered:
                errors = list(errors) + [{'reason': '仍有填写位置未关联源字段；请按 mapping_requests 一次补齐或说明留空原因',
                                          'count': len(uncovered)}]
            record['plan'] = plan
            if errors:
                record.update(status='needs_mapping', error=errors)
                task['issues'].append({'template': name, 'problems': errors,
                                      'current_positions': [{'field': r['field'], 'target': r['target']}
                                                            for r in store.rules_for(path)]})
            else:
                signature = review.execution_signature(plan)
                if prior.get('status') == 'running' and prior.get('signature') == signature:
                    run = store.conn.execute("SELECT status FROM fill_op WHERE run_id=? AND kind='run'", (plan['run_id'],)).fetchone()
                    artifact = store.conn.execute("SELECT artifact_path,artifact_sha256 FROM fill_op WHERE run_id=? AND kind='artifact' AND opened_ok=1 ORDER BY id DESC LIMIT 1", (plan['run_id'],)).fetchone()
                    if run and run[0] == 'ok' and artifact:
                        recovered = Path(artifact[0])
                        if recovered.is_file() and recovered.is_relative_to(work) and sha256_file(recovered) == artifact[1]:
                            prior.update(status='completed', output=recovered.relative_to(work).as_posix(), output_hash=artifact[1])
                            record.update(prior)
                output = work / prior.get('output', '__absent__')
                if prior.get('status') == 'completed' and prior.get('signature') == signature and output.is_file() and sha256_file(output) == prior.get('output_hash'):
                    record['status'] = 'completed'
                else:
                    record['status'] = 'pending'
                    record.pop('error', None)
                record['signature'] = signature
        except Exception as exc:
            record.update(status='needs_mapping', error=str(exc))
            task['issues'].append({'template': name, 'reason': str(exc)})
        task['documents'][name] = record


def advance(store, work, task, *, execute=True):
    questions = source.prepare(store, work, task)
    if questions:
        task.update(status='awaiting_source', questions=questions)
        save(store, task)
        return
    if not task.get('rows'):
        task['questions'] = []
        save(store, task)
        return
    if not task.get('source_ready'):
        source.ingest(store, work, task)
        save(store, task)
    prepare_documents(store, work, task)
    if any(r.get('status') == 'needs_mapping' for r in task['documents'].values()):
        task.update(status='needs_mapping', questions=[])
        save(store, task)
        return
    task['questions'] = review.questions(task)
    if task['questions']:
        task['status'] = 'awaiting_fill'
    elif execute:
        execute_documents(store, work, task)
    save(store, task)


def execute_documents(store, work, task):
    if any(r.get('status') == 'needs_mapping' for r in task['documents'].values()):
        task.update(status='needs_mapping', questions=[])
        return
    for name, record in task['documents'].items():
        if record['status'] != 'pending':
            continue
        plan = record['plan']
        plan['run_id'] = f"{task['batch']}-V2-{uuid.uuid4().hex[:12]}"
        p = work / 'work' / task['batch'] / f"v2-{task['id']}-{state.digest(name)[:12]}.json"
        p.parent.mkdir(parents=True, exist_ok=True)
        p.write_text(json.dumps(plan, ensure_ascii=False), encoding='utf-8')
        record['status'] = 'running'
        save(store, task)
        try:
            if not plan['rows'] and record.get('slots') and not record.get('unmapped_slots'):
                # Every discovered position is protected or expressly left blank.
                # This is a documented unchanged copy, not a fictitious fill run.
                import shutil
                target = work / 'out' / task['batch'] / '留空' / (Path(name).stem + '_留空副本' + Path(name).suffix)
                target.parent.mkdir(parents=True, exist_ok=True)
                if target.exists():
                    target = target.with_stem(target.stem + '-' + uuid.uuid4().hex[:8])
                shutil.copy2(state.inside(work, name), target)
                record.update(status='completed', output=target.relative_to(work).as_posix(),
                              output_hash=sha256_file(target), signature=review.execution_signature(plan),
                              filled=0, failed=0)
                save(store, task)
                continue
            env = {**os.environ, 'PYTHONIOENCODING': 'utf-8', 'PYTHONUTF8': '1'}
            cp = subprocess.run([sys.executable, '-m', 'office_kit', 'db-fill', str(state.inside(work, name)),
                                 '--work', str(work), '--batch', task['batch'], '--plan', str(p), '--confirmed'],
                                cwd=ROOT, env=env, encoding='utf-8', capture_output=True, timeout=120)
            task['counts']['fill_processes'] += 1
            result = json.loads(cp.stdout)
            if cp.returncode or not result.get('ok'):
                raise ValueError(result.get('error') or cp.stderr[-500:])
            data = result['data']
            actual = data.get('templates', [])
            if data.get('blocked') or not actual or not all(r.get('deliverable') for r in actual):
                if not data.get('blocked') and not actual:
                    import shutil
                    target = work / 'out' / task['batch'] / '留空' / (Path(name).stem + '_已填写' + Path(name).suffix)
                    target.parent.mkdir(parents=True, exist_ok=True)
                    if target.exists():
                        target = target.with_stem(target.stem + '-' + uuid.uuid4().hex[:8])
                    shutil.copy2(state.inside(work, name), target)
                    actual = [{'output': str(target), 'filled': 0, 'failed': 0}]
                else:
                    raise ValueError(json.dumps(data, ensure_ascii=False)[:1200])
            output = Path(actual[0]['output']).resolve()
            record.update(status='completed', output=output.relative_to(work).as_posix(),
                          output_hash=sha256_file(output), signature=review.execution_signature(plan),
                          filled=actual[0].get('filled', 0), failed=actual[0].get('failed', 0))
        except Exception as exc:
            record.update(status='failed', error=str(exc))
        save(store, task)
    states = [r['status'] for r in task['documents'].values()]
    task['status'] = ('completed' if states and all(s == 'completed' for s in states) else
                      'partial' if 'completed' in states else 'needs_mapping' if 'needs_mapping' in states else 'failed')
    if task['status'] == 'completed' and task.get('source_issues'):
        task['status'] = 'partial'
    task['questions'] = []
    from .report import write
    write(work, task)


def update_positions(store, work, task, updates):
    if not isinstance(updates, list) or not updates:
        raise ValueError('没有提供需要修改的位置')
    checked, blanks = mapping.compile_updates(store, work, task, updates)
    # Check legacy replacements too before saving any change in this request.
    for path, update in checked:
        mapping.combine(store, path, update)
    for path, update in checked:
        name = path.relative_to(work).as_posix()
        tid = store.register_template(path, task['batch'])
        if update.get('slot_id'):
            mapping.retire_overlap(store, path, update['target'], update['field'], task['batch'])
            task['documents'][name].setdefault('blank_slots', {}).pop(update['slot_id'], None)
        target = mapping.combine(store, path, update)
        store.add_rule(tid, update['field'], update.get('label', update['field']), target,
                       confidence=1.0, decided_by='model', batch_no=task['batch'])
    for name, slot_id, reason in blanks:
        record = task['documents'][name]
        slot = next(s for s in record['slots'] if s['id'] == slot_id)
        mapping.retire_overlap(store, work / name, slot['target'], None, task['batch'])
        record.setdefault('blank_slots', {})[slot_id] = reason
    save(store, task)
    advance(store, work, task)


def dispatch(data):
    if not isinstance(data, dict) or not isinstance(data.get('work'), str):
        raise ValueError('请求需要明确 work 工作区')
    work = Path(data['work']).resolve(strict=True)
    if not work.is_dir():
        raise ValueError('工作区必须是目录')
    if not is_workroot(work):
        if data.get('action') != 'start':
            raise ValueError('工作区尚未初始化，请先开始任务')
        init_workroot(work)
    with state.work_lock(work), StoreV2(work / 'db/workflow.db', actor='workflow') as store:
        state.initialise(store.conn)
        action = data.get('action')
        task = None
        try:
            if action == 'start':
                task = start(store, work, data)
                advance(store, work, task)
            else:
                task = state.get(store.conn, 'office_v2_task', data.get('task_id', ''))
                if task is None:
                    raise ValueError('找不到该任务')
                rebase(store, work, task)
                if action == 'update_positions':
                    update_positions(store, work, task, data.get('updates'))
                elif action in ('status', 'resume'):
                    previous_questions = task.get('questions', [])
                    previous_status = task.get('status')
                    advance(store, work, task, execute=False)
                    changed = state.digest(previous_questions) != state.digest(task.get('questions', []))
                    if action == 'resume' and not changed and review.validate(task.get('questions', []), data.get('answers', [])):
                        if previous_status == 'awaiting_source':
                            source.save_answers(store, task, data['answers'])
                        else:
                            review.save(task, data['answers'])
                        advance(store, work, task)
                    elif action == 'status' and not task.get('questions'):
                        execute_documents(store, work, task)
                    save(store, task)
                else:
                    raise ValueError('不支持的任务操作')
        except Exception as exc:
            if task is None:
                raise
            result = response(work, task)
            result.update(ok=False, status='failed', questions=[], error=str(exc),
                          resume_status=task.get('status'),
                          recovery='修复后使用此 task_id 恢复原任务；已入库答案保留，不要更换批次或拆成新任务。')
            return result
        return response(work, task)
