"""Deterministic fill-task entry. Model calls one tool; code advances it."""
from __future__ import annotations
import json
import os
import re
from pathlib import Path
import subprocess
import sys
import uuid
from office_kit.store_v2 import StoreV2, allocate_batch, sha256_file, valid_batch
from office_kit.common import resolve_inputs
from office_kit.workroot import init_workroot, is_workroot
from office_kit.cli import build_parser, _dispatch
from office_kit.harness import build_fill_plan, current_docx_target_errors, current_xlsx_gaps
from office_kit.target_validation import template_read_cache
from . import storage as state, source, review, mapping, mapping_view, timing

ROOT = Path(__file__).resolve().parents[1]
WorkflowError = ValueError
POSITION_PARSE_VERSION = 6
SUBJECT_PLAN_VERSION = 17
# Bump when document planning itself changes.  Equal inputs may then reuse a
# stored plan without re-reading every template's slots.
DOCUMENT_PLAN_VERSION = 1
XLSX_WRITER_VERSION = 1
# Kept here because runner owns plan reuse.  Bump alongside a mapping semantic
# change that can alter automatic positions without changing template bytes.
MAPPING_COMPATIBILITY_VERSION = 3


def command(argv):
    return json.loads(_dispatch(build_parser().parse_args(argv)).to_json())


def save(store, task):
    state.put(store.conn, 'office_v2_task', task['id'], task)


def response(work, task, execution_summary=None):
    results = []
    for name, record in task.get('documents', {}).items():
        result = {k: record[k] for k in ('status', 'error', 'filled', 'failed', 'coverage', 'blank_slots') if k in record}
        result['template'] = name
        if record.get('output') and record.get('status') == 'completed':
            result['output'] = str(work / record['output'])
        results.append(result)
    needs_mapping = task['status'] == 'needs_mapping'
    if needs_mapping:
        # Detailed errors and rules remain readable through mapping_read, once only.
        for item in results:
            if item.get('error'):
                item['error'] = '待整理位置；原因见 mapping_read 的 issues 页'
            item.pop('blank_slots', None)
        results = results[:20]
    issues = task.get('issues', [])
    if needs_mapping and mapping_view.size(issues) > 1800:
        issues = [{'count': len(issues), 'read': {'section': 'issues'}, 'note': '完整问题按页读取'}]
    result = {'ok': task['status'] != 'failed', 'status': task['status'], 'task_id': task['id'],
            'work': str(work), 'batch': task['batch'], 'questions': task.get('questions', []),
            'issues': issues, 'results': results, 'counters': task['counts'],
            'timing': timing.summary(task),
            'report': str(work / task['report']) if task.get('report') and task['status'] in ('completed', 'partial') else None,
            'mapping_requests': [], 'available_fields': [],
            'mapping_revision': mapping_view.revision(task),
            'unsigned': True, 'note': ('部分源材料尚未解析，请查看 issues 中的文件和原文；已生成文件仅包含已处理内容。'
                                     if task.get('source_issues') else '')
                                    + '填报完成的文件供用户复核；未代用户签核。'}
    waits = task.get('timing', {}).get('stages', {}).get('user_confirmation_wait', {})
    if waits.get('count'):
        result['interaction_summary'] = {
            'recorded_question_batches': waits['count'],
            'wait_seconds': waits['seconds'],
            'cancelled_batches': task.get('timing', {}).get('cancelled_waits', 0),
            'pending_questions': len(task.get('questions', [])),
            'note': '这是工具内部原生问题面板的历史记录；questions 为空只表示当前无待答，不能据此声称从未提问。'}
    if task.get('rejected_updates'):
        result['rejected_updates'] = task['rejected_updates']
        # Existing completed artifacts remain usable, but the latest update
        # request was not fully accepted and must not look like controller
        # success.
        result['ok'] = False
        result['success_with_rejected'] = True
    if needs_mapping:
        result.update(mapping_view.initial(task))
    if task.get('delivery') and not needs_mapping:
        result['delivery'] = task['delivery']
    if execution_summary is not None:
        result['execution_summary'] = execution_summary
        if task['status'] == 'completed' and result['ok']:
            result['next_action'] = (
                '本次实际执行统计见 execution_summary；当前产物已逐文件核验路径和 SHA-256。'
                '按 delivery.present_calls 交付并回复统计即可。'
                '输出目录保留旧版本，遍历目录不能判断本次生成数量。')
    return result


def execution_summary(before, task):
    """Describe only this dispatch call; never persist or scan directories."""
    generated, reused = [], []
    for name, record in task.get('documents', {}).items():
        old = before.get('documents', {}).get(name, {})
        if record.get('status') != 'completed' or not record.get('output'):
            continue
        # A new run id is the durable proof that this call executed or copied
        # the template.  Recovery can restore a missing output path for an
        # already successful run, which must remain a reuse.
        changed = (record.get('run_id') != old.get('run_id') if record.get('run_id')
                   else (record.get('output') != old.get('output') or
                         record.get('output_hash') != old.get('output_hash')))
        (generated if changed else reused).append(name)
    return {'generated_files': len(generated), 'reused_files': len(reused),
            'generated_templates': generated}


def preserve_output(record):
    if not record.get('output'):
        return
    history = record.setdefault('output_history', [])
    if not any(item.get('output') == record['output'] for item in history):
        history.append({key: record.get(key) for key in ('output', 'output_hash', 'run_id')})


def refresh_delivery(store, work, task):
    from .delivery import build_delivery
    artifacts = [dict(r) for r in store.conn.execute(
        "SELECT a.*,t.path AS template_path,t.name AS template_name FROM fill_op a "
        "LEFT JOIN template t ON t.id=a.template_id WHERE a.kind='artifact' AND a.batch_no=? ORDER BY a.id",
        (task['batch'],))]
    report = work / 'out' / task['batch'] / '_报告' / f"toolV2-{task['id']}-全部文件.md"
    task['delivery'] = build_delivery(task, artifacts, str(report) if task['status'] in ('completed', 'partial') else None)


def rebase(store, work, task):
    old = task.get('last_work')
    if not old or old == str(work):
        return
    for table, column in [('template', 'path'), ('source', 'path'), ('source', 'copy_path'),
                          ('fill_op', 'artifact_path')]:
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
            raise ValueError(f'{key} 必须是非空文件或目录路径数组')
        expanded = []
        for name in names:
            # Check the supplied root before expansion: resolve_inputs accepts
            # directories, but a directory (or symlink) outside this work area
            # must never be enumerated as a source of office materials.
            supplied = (work / name).resolve(strict=True)
            if not supplied.is_relative_to(work):
                raise ValueError(f'文件不在本工作区内：{name}')
            files = resolve_inputs(str(supplied))
            if key == 'targets':
                # A target directory may contain notes and Office lock files.
                # Only real Word/Excel templates participate in the task.
                files = [p for p in files if p.suffix.lower() in ('.docx', '.xlsx')
                         and not p.name.startswith('~$')]
                if supplied.is_file() and not files:
                    raise ValueError('当前填报入口支持 Word .docx 与 Excel .xlsx 目标模板')
            expanded.extend(files)
        normalized = {state.inside(work, p).relative_to(work).as_posix() for p in expanded}
        if not normalized:
            raise ValueError(f'{key} 未找到可用文件')
        lists.append(sorted(normalized))
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
            if prior.get('subject_plan_version') != SUBJECT_PLAN_VERSION:
                # Reconsider only old machine exclusions caused by the removed
                # document-wide ownership gate. Missing-data/user blanks remain.
                record['blank_slots'] = {k: v for k, v in record.get('blank_slots', {}).items()
                                         if '跨主体红线' not in str(v)
                                         and '不属于本产物' not in str(v)}
            record['subject_plan_version'] = SUBJECT_PLAN_VERSION
            source_changed = (prior.get('blank_source_content_signature') is not None and
                              prior.get('blank_source_content_signature') !=
                              task.get('source_content_signature'))
            # A user explicitly left this position blank.  New source content
            # does not revoke that decision; only a changed template invalidates
            # the physical slot it referred to.
            if record.get('blank_template_hash') != sha256_file(path):
                record['blank_slots'] = {}
            record['blank_source_signature'] = task.get('source_signature')
            record['blank_source_content_signature'] = task.get('source_content_signature')
            template_hash = sha256_file(path)
            record['blank_template_hash'] = template_hash
            if (source_changed or prior.get('subject_plan_version') != SUBJECT_PLAN_VERSION) and record.get('blank_slots'):
                # “来源未提供” is a missing-data deferral, not a permanent
                # user refusal.  Reopen only that slot when a newly supplied
                # source fact is a unique high-confidence candidate.  Other
                # explicit leave-blank reasons remain untouched.
                by_id = {slot['id']: slot for slot in record.get('slots', [])}
                for slot_id, reason in list(record['blank_slots'].items()):
                    if not re.search(r'(?:来源|源文件|材料|资料).{0,80}(?:未提供|未给出|没有|未记载|缺失)|缺失', str(reason)):
                        continue
                    slot = by_id.get(slot_id)
                    if slot is None:
                        continue
                    candidates = mapping.scoped_candidates(store, path, slot, catalog)
                    if (candidates and candidates[0]['score'] >= .85 and
                            (len(candidates) == 1 or
                             candidates[1]['score'] != candidates[0]['score'])):
                        from office_kit.target_validation import value_target_issue
                        from office_kit.harness import _missing_source_value_reason
                        field = candidates[0]['field']
                        value = catalog[field].get('value')
                        if (value not in (None, '') and
                                not _missing_source_value_reason(field, value) and
                                not value_target_issue(path, slot['target'], str(value))):
                            record['blank_slots'].pop(slot_id, None)
            inputs = state.digest([DOCUMENT_PLAN_VERSION, MAPPING_COMPATIBILITY_VERSION,
                                   mapping.SLOT_DISCOVERY_VERSION, POSITION_PARSE_VERSION,
                                   SUBJECT_PLAN_VERSION, template_hash,
                                   store.rules_for(path), facts, roles,
                                   record.get('blank_slots'),
                                   *([XLSX_WRITER_VERSION] if path.suffix.lower() == '.xlsx' else [])])
            # A status/read call and an idempotent position update must not
            # rescan every Word/Excel slot when all inputs are unchanged.
            # Source/template/rule/fact/role/version changes are all in the
            # digest, so only affected documents fall through to planning.
            output_ok = True
            if prior.get('status') == 'completed':
                output = work / prior.get('output', '__absent__')
                output_ok = (output.is_file() and bool(prior.get('output_hash')) and
                             sha256_file(output) == prior.get('output_hash'))
            if prior.get('inputs') == inputs and prior.get('plan') and output_ok:
                record.update(inputs=inputs)
                if record.get('status') == 'needs_mapping' and record.get('validation_errors'):
                    task['issues'].append({
                        'template': name,
                        'problems': record.get('validation_errors', record.get('error', [])),
                        'current_positions': record.get('current_positions', []),
                    })
                task['documents'][name] = record
                continue
            if (prior.get('proposed_version') != POSITION_PARSE_VERSION
                    or prior.get('proposed_hash') != sha256_file(path)):
                command(['db-propose', str(path), '--work', str(work), '--batch', task['batch']])
                record['proposed_hash'] = sha256_file(path)
                record['proposed_version'] = POSITION_PARSE_VERSION
                task['counts']['position_proposals'] += 1
            mapping.auto_map(store, path, record, catalog, task['batch'])
            uncovered = mapping.inspect(store, path, record, catalog)
            record['current_positions'] = [{'field': r['field'], 'target': r['target']} for r in store.rules_for(path)]
            # `auto_map` may add or retire rules.  Persist the final rule
            # state, not the pre-migration state used for the fast-path gate;
            # otherwise the next unchanged status unnecessarily re-previews.
            inputs = state.digest([DOCUMENT_PLAN_VERSION, MAPPING_COMPATIBILITY_VERSION,
                                   mapping.SLOT_DISCOVERY_VERSION, POSITION_PARSE_VERSION,
                                   SUBJECT_PLAN_VERSION, template_hash,
                                   store.rules_for(path), facts, roles,
                                   record.get('blank_slots'),
                                   *([XLSX_WRITER_VERSION] if path.suffix.lower() == '.xlsx' else [])])
            if prior.get('inputs') == inputs and prior.get('plan'):
                plan = prior['plan']
            else:
                plan = build_fill_plan(store, [path], batch_no=task['batch'],
                                       run_id=prior.get('plan', {}).get('run_id') or f"{task['batch']}-V2-{uuid.uuid4().hex[:12]}")
                task['counts']['previews'] += 1
            plan.update(db_path=str(work / 'db/workflow.db'), template_files=[str(path)])
            if path.suffix.lower() == '.xlsx':
                plan['writer_version'] = XLSX_WRITER_VERSION
            if prior.get('plan') and prior.get('inputs') != inputs:
                from office_kit.fill_decisions import reuse_unchanged_choices
                reuse_unchanged_choices(prior['plan'], plan)
            plan['position_inventory'] = {'template_hash': sha256_file(path),
                                          'blank_targets': [s['target'] for s in record.get('slots', [])
                                                            if s['id'] in record.get('blank_slots', {})]}
            if prior.get('inputs') == inputs:
                validation_errors = prior.get('validation_errors', [])
            else:
                # A plain unmapped slot is represented by coverage and
                # mapping_requests.  It is not a validation error and should
                # not duplicate every position into the issues page.
                validation_errors = current_docx_target_errors([path], plan)
                validation_errors += current_xlsx_gaps(store, [path], plan)
                for rule in store.rules_for(path):
                    try:
                        mapping.validate_target(path, rule['target'])
                    except Exception as exc:
                        validation_errors.append({'field': rule['field'], 'reason': str(exc),
                                                  'target': rule['target']})
            if record.get('slots') and not uncovered and not store.rules_for(path):
                # User requested missing/uncertain data remain blank. A template
                # whose positions were all accounted for can be copied unchanged.
                plan['rows'] = []
                validation_errors = []
            record.update(inputs=inputs, validation_errors=list(validation_errors))
            record['plan'] = plan
            if validation_errors:
                preserve_output(record)
                record.update(status='needs_mapping', error=validation_errors)
                task['issues'].append({'template': name, 'problems': validation_errors,
                                      'current_positions': [{'field': r['field'], 'target': r['target']}
                                                            for r in store.rules_for(path)]})
            elif uncovered:
                preserve_output(record)
                record.update(status='needs_mapping')
                record.pop('error', None)
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
                    preserve_output(record)
                    record['status'] = 'pending'
                    record.pop('error', None)
                record['signature'] = signature
        except Exception as exc:
            preserve_output(record)
            record.update(status='needs_mapping', error=str(exc))
            task['issues'].append({'template': name, 'reason': str(exc)})
        task['documents'][name] = record


def advance(store, work, task, *, execute=True):
    with timing.measure(task, 'source_parse_and_check'):
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
        with timing.measure(task, 'source_ingest'):
            source.ingest(store, work, task)
        save(store, task)
    with timing.measure(task, 'position_plan_and_check'):
        prepare_documents(store, work, task)
    if any(r.get('status') == 'needs_mapping' for r in task['documents'].values()):
        task.update(status='needs_mapping', questions=[])
        save(store, task)
        return
    task['questions'] = review.questions(task)
    if task['questions']:
        task['status'] = 'awaiting_fill'
    elif execute:
        with timing.measure(task, 'generation_and_check'):
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
        preserve_output(record)
        plan['run_id'] = f"{task['batch']}-V2-{uuid.uuid4().hex[:12]}"
        record['run_id'] = plan['run_id']
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
    refresh_delivery(store, work, task)
    from .report import write
    write(work, task)


def update_positions(store, work, task, updates):
    if not isinstance(updates, list) or not updates:
        raise ValueError('没有提供需要修改的位置')
    grouped = {}
    for update in updates:
        name = update.get('template') if isinstance(update, dict) else None
        grouped.setdefault(name, []).append(update)
    rejected = []
    accepted = False
    task['rejected_updates'] = []
    for name, template_updates in grouped.items():
        # Keep a template atomic, but do not make one bad template discard
        # correctly scoped changes for other templates in the same request.
        try:
            checked, blanks = mapping.compile_updates(store, work, task, template_updates)
            # Validate all legacy replacements before this template writes.
            for path, update in checked:
                mapping.combine(store, path, update)
        except (ValueError, KeyError, TypeError) as exc:
            rejected_item = {'template': name, 'updates': template_updates,
                             'reason': str(exc)}
            slot_id = getattr(exc, 'slot_id', None)
            if slot_id and name in task.get('documents', {}):
                slot = next((item for item in task['documents'][name].get('slots', [])
                             if item['id'] == slot_id), None)
                rejected_item.update(slot_id=slot_id, label=getattr(exc, 'label', None) or (slot or {}).get('label'),
                                     positions_read={'section': 'positions', 'template': name})
                if slot is None:
                    import difflib
                    slots = task['documents'][name].get('slots', [])
                    nearby = difflib.get_close_matches(slot_id, [s['id'] for s in slots], n=3, cutoff=.85)
                    rejected_item['similar_positions'] = [
                        {'id': s['id'], 'label': s.get('label')} for key in nearby for s in slots if s['id'] == key]
            rejected.append(rejected_item)
            continue
        for path, update in checked:
            resolved_name = path.relative_to(work).as_posix()
            tid = store.register_template(path, task['batch'])
            if update.get('slot_id'):
                mapping.retire_overlap(store, path, update['target'], update['field'], task['batch'])
                task['documents'][resolved_name].setdefault('blank_slots', {}).pop(update['slot_id'], None)
            target = mapping.combine(store, path, update)
            store.add_rule(tid, update['field'], update.get('label', update['field']), target,
                           confidence=1.0, decided_by='model', batch_no=task['batch'])
        for resolved_name, slot_id, reason in blanks:
            record = task['documents'][resolved_name]
            slot = next(s for s in record['slots'] if s['id'] == slot_id)
            mapping.retire_overlap(store, work / resolved_name, slot['target'], None, task['batch'])
            record.setdefault('blank_slots', {})[slot_id] = reason
        accepted = True
    task['rejected_updates'] = rejected
    save(store, task)
    if accepted:
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
    if data.get('action') == 'read_mapping':
        # Readers use one committed task snapshot. They neither initialise the
        # schema nor take the writer lock, so independent evidence reads do not
        # block each other or create another fill execution.
        import sqlite3
        with sqlite3.connect((work / 'db/workflow.db').as_uri() + '?mode=ro', uri=True) as conn:
            task = state.get(conn, 'office_v2_task', data.get('task_id', ''))
        if task is None:
            raise ValueError('找不到该任务')
        task['last_work'] = str(work)
        result = {'ok': True, 'status': task['status'], 'task_id': task['id'],
                  'work': str(work), 'batch': task['batch'], 'counters': task['counts']}
        with template_read_cache():
            try:
                result['mapping_page'] = mapping_view.page(task, data.get('mapping_read') or {})
            except Exception as exc:
                result.update(ok=False, status='failed', error=str(exc), questions=[],
                              resume_status=task['status'])
        return result
    with state.work_lock(work), template_read_cache(), StoreV2(work / 'db/workflow.db', actor='workflow') as store:
        state.initialise(store.conn)
        action = data.get('action')
        task = None
        started = None
        try:
            if action == 'start':
                task = start(store, work, data)
                before_execution = {'documents': {name: dict(record) for name, record in task.get('documents', {}).items()}}
                started = timing.begin(task)
                advance(store, work, task)
            else:
                task = state.get(store.conn, 'office_v2_task', data.get('task_id', ''))
                if task is None:
                    raise ValueError('找不到该任务')
                rebase(store, work, task)
                before_execution = {'documents': {name: dict(record) for name, record in task.get('documents', {}).items()}}
                if action == 'read_mapping':
                    return {'ok': True, 'status': task['status'], 'task_id': task['id'],
                            'work': str(work), 'batch': task['batch'],
                            'mapping_page': mapping_view.page(task, data.get('mapping_read') or {}),
                            'counters': task['counts']}
                started = timing.begin(task, data.get('user_wait'))
                if action == 'record_wait':
                    pass
                elif action == 'update_positions':
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
                        with timing.measure(task, 'generation_and_check'):
                            execute_documents(store, work, task)
                    save(store, task)
                else:
                    raise ValueError('不支持的任务操作')
        except Exception as exc:
            if task is None:
                raise
            if started is not None:
                timing.end(task, started)
                save(store, task)
            result = ({'task_id': task['id'], 'work': str(work), 'batch': task['batch']}
                      if action == 'read_mapping' else response(work, task))
            result.update(ok=False, status='failed', questions=[], error=str(exc),
                          resume_status=task.get('status'),
                          recovery='修复后使用此 task_id 恢复原任务；已入库答案保留，不要更换批次或拆成新任务。')
            return result
        if started is not None:
            timing.end(task, started)
            refresh_delivery(store, work, task)
            if task['status'] in ('completed', 'partial'):
                from .report import write
                write(work, task)
            save(store, task)
        return response(work, task, execution_summary(before_execution, task))
