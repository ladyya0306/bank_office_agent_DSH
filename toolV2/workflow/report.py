"""One complete task report, alongside the existing per-file engine reports."""
from pathlib import Path
from .review import location
from .delivery import artifact_versions
from office_kit.fill_decisions import saved_choices


def version_lines(task, artifact_records=None):
    versions = (task.get('delivery', {}).get('versions', []) if artifact_records is None
                else artifact_versions(task, artifact_records))
    lines = ['## 结果文件版本', '', '当前版本可用于本次交付；被替代版本仍保留在原位置供追溯。', '']
    if not versions:
        lines += ['尚无可列出的登记产物版本。', '']
    else:
        for item in versions:
            label = {'current': '当前版本', 'superseded': '已被替代，保留',
                     'previous': '旧版待更新，新版尚未成功',
                     'historical': '历史产物'}[item['status']]
            lines.append(f"- {label}：{item['template']}；文件 `{item['path']}`；运行 `{item.get('run_id') or '未知'}`")
        lines.append('')
    return lines


def timing_lines(task):
    timing = task.get('timing') or {}
    stages = timing.get('stages') or {}
    labels = {
        'tool_calls': '工具调用总跨度（包含程序子阶段，不与下列阶段相加）',
        'source_parse_and_check': '来源解析与检查',
        'source_ingest': '来源事实入库',
        'position_plan_and_check': '位置整理与检查',
        'user_confirmation_wait': '用户确认等待（仅记录到原生问题的实测时间）',
        'generation_and_check': '生成与检查',
        'mapping_interval_unclassified': '位置整理调用间隔（未细分）',
        'other_interval_unclassified': '其它调用间隔（未细分）',
    }
    entries = []
    for name, data in stages.items():
        if not isinstance(data, dict) or data.get('seconds') is None:
            continue
        label = labels.get(name, name)
        count = data.get('count')
        suffix = f"；次数 {count}" if count is not None else ''
        entries.append(f"- {label}：{data['seconds']} 秒{suffix}")
    if not entries:
        return []
    lines = ['## 阶段用时记录', '', *entries]
    note = timing.get('summary', {}).get('note') if isinstance(timing.get('summary'), dict) else timing.get('note')
    if not note and stages:
        from .timing import summary
        note = summary(task).get('note')
    if note:
        lines.append(f'说明：{note}')
    lines.append('')
    return lines


def write(work, task, artifact_records=None):
    lines = ['# 本次填报结果', '', f"业务批次：{task['batch']}", '',
             '这是填写结果说明，不是签核结论。原始资料和目标模板没有被覆盖。', '',
             '程序识别的空位不是所有视觉空白的保证；文件仍需按业务内容复核。', '']
    lines += version_lines(task, artifact_records)
    lines += timing_lines(task)
    for name, record in task['documents'].items():
        lines += [f'## {Path(name).name}', '',
                  f"状态：{record.get('status')}；实际填写 {record.get('filled', 0)} 处。", '']
        if record.get('status') == 'completed' and record.get('output'):
            lines += [f"结果文件：{record['output']}", '']
        if record.get('error'):
            lines += [f"未完成原因：{record['error']}", '']
        slots = {s['id']: s for s in record.get('slots', [])}
        for key, reason in record.get('blank_slots', {}).items():
            slot = slots.get(key)
            if slot:
                where = location({'template': name, 'target': slot['target'], 'label': slot['label']})
                lines.append(f'- 留空：{where}。原因：{reason}')
        for slot in slots.values():
            if slot.get('protected'):
                where = location({'template': name, 'target': slot['target'], 'label': slot['label']})
                lines.append(f"- 未代签核：{where}。{slot.get('protected_reason', '')}")
        plan = record.get('plan') or {}
        try:
            choices, _ = saved_choices(plan) if plan.get('db_path') else ({'blank': []}, [])
            blank_numbers = {int(n) for n in choices.get('blank', [])}
        except Exception as exc:
            blank_numbers = set()
            lines.append(f'- 留空确认记录读取失败：{exc}。本报告无法核对这份文件的留空选择。')
        for row in plan.get('rows', []):
            if row.get('kind') != 'slot' or int(row.get('n', -1)) not in blank_numbers:
                continue
            where = location({'template': name, 'target': row['target'], 'label': row.get('label', row['field'])})
            reason = row.get('local_issue') or ('来源没有建议值，用户确认留空'
                                                 if row.get('value') in (None, '')
                                                 else '用户确认留空')
            lines.append(f'- 用户确认留空：{where}。原因：{reason}')
        for row in record.get('plan', {}).get('rows', []):
            if (row.get('kind') == 'slot' and row.get('value') in (None, '')
                    and int(row.get('n', -1)) not in blank_numbers):
                lines.append(f"- 来源未给出建议值：{row.get('field')}，以用户确认结果为准。")
        lines.append('')
    path = work / 'out' / task['batch'] / '_报告' / f"toolV2-{task['id']}-全部文件.md"
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text('\n'.join(lines), encoding='utf-8')
    task['report'] = path.relative_to(work).as_posix()
