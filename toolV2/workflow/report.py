"""One complete task report, alongside the existing per-file engine reports."""
from pathlib import Path
from .review import location


def write(work, task):
    lines = ['# 本次填报结果', '', f"业务批次：{task['batch']}", '',
             '这是填写结果说明，不是签核结论。原始资料和目标模板没有被覆盖。', '',
             '程序识别的空位不是所有视觉空白的保证；文件仍需按业务内容复核。', '']
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
        for row in record.get('plan', {}).get('rows', []):
            if row.get('kind') == 'slot' and row.get('value') in (None, ''):
                lines.append(f"- 来源未给出建议值：{row.get('field')}，以用户确认结果为准。")
        lines.append('')
    path = work / 'out' / task['batch'] / '_报告' / f"toolV2-{task['id']}-全部文件.md"
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text('\n'.join(lines), encoding='utf-8')
    task['report'] = path.relative_to(work).as_posix()
