"""Connect this portable tool folder to an existing local DSH installation."""
from __future__ import annotations

import argparse
import importlib
import json
import os
from pathlib import Path
import re
import shutil
import sys
from datetime import datetime

ROOT = Path(__file__).resolve().parent
START = '# BEGIN OFFICE TOOL V2'
END = '# END OFFICE TOOL V2'


def install(dsh_root: Path) -> dict:
    dsh = dsh_root.resolve(strict=True)
    workspace = dsh / 'workspace'
    profile = workspace / 'profiles/web/cordis.patch.yml'
    hooks_path = workspace / 'hooks.json'
    plugin = workspace / 'plugins/office-tool-v2'
    skill = workspace / 'skills/office-suite'
    sources = [ROOT / 'office.py', ROOT / 'dsh-plugin/index.mjs',
               ROOT / 'skills/office-suite-v2/SKILL.md', ROOT / 'hooks/bank_gate.py']
    for p in sources + [profile, hooks_path]:
        if not p.is_file():
            raise RuntimeError(f'缺少文件：{p}')
    for name in ('docx', 'openpyxl', 'lxml'):
        importlib.import_module(name)
    hooks = json.loads(hooks_path.read_text(encoding='utf-8-sig'))
    command = f'"{sys.executable}" "{ROOT / "hooks/bank_gate.py"}"'
    for event in ('SessionStart', 'UserPromptSubmit'):
        changed = False
        for group in hooks.get('hooks', {}).get(event, []):
            for hook in group.get('hooks', []):
                if hook.get('type') == 'command' and 'bank_gate.py' in hook.get('command', ''):
                    hook['command'] = command
                    changed = True
        if not changed:
            raise RuntimeError(f'现有 {event} 配置没有 bank_gate.py，未修改配置')
    raw = profile.read_text(encoding='utf-8-sig')
    raw = re.sub(re.escape(START) + r'.*?' + re.escape(END) + r'\s*', '', raw, flags=re.S)
    # Existing local plugin paths are rebased after an offline move.
    raw = re.sub(r'''(?m)^(\s*configPath:)\s*['"].*hooks\.json['"]\s*$''',
                 lambda m: m[1] + ' ' + json.dumps(str(hooks_path), ensure_ascii=False), raw)
    for name, relative in [('bank-approval', 'bank-approval/index.mjs'),
                           ('bank-approval-ui', 'bank-approval-ui/lib/index.js'),
                           ('office-fill-review', 'office-fill-review/index.mjs')]:
        raw = re.sub(r'(?m)(^\s*- id: ' + name + r'\s*\n\s*name:)\s*[^\r\n]+',
                     lambda m, rel=relative: m[1] + ' ' + json.dumps(str(workspace / 'plugins' / rel)), raw)
    config = '\n'.join([
        START, *(['- id: office-fill-review', '  disabled: true'] if re.search(r'(?m)^\s*- id: office-fill-review\s*$', raw) else []), '- insert:',
        '    - id: office-tool-v2',
        '      name: ' + json.dumps(str(plugin / 'index.mjs'), ensure_ascii=False),
        '      config:', '        toolRoot: ' + json.dumps(str(ROOT), ensure_ascii=False),
        '        python: ' + json.dumps(sys.executable, ensure_ascii=False), END, '',
    ])
    wanted_profile = raw.rstrip() + '\n\n' + config
    wanted_skill = (ROOT / 'skills/office-suite-v2/SKILL.md').read_text(encoding='utf-8')
    wanted_skill = re.sub(r'(?m)^name:\s*office-suite-v2\s*$', 'name: office-suite', wanted_skill)
    plugin_current = plugin.is_dir() and all(
        (plugin / p.name).is_file() and (plugin / p.name).read_bytes() == p.read_bytes()
        for p in (ROOT / 'dsh-plugin').glob('*.mjs'))
    if (profile.read_text(encoding='utf-8-sig') == wanted_profile
            and json.loads(hooks_path.read_text(encoding='utf-8-sig')) == hooks
            and (skill / 'SKILL.md').is_file()
            and (skill / 'SKILL.md').read_text(encoding='utf-8') == wanted_skill
            and plugin_current):
        return {'ok': True, 'changed': False, 'tool_root': str(ROOT),
                'message': 'toolV2 接入已是当前版本，无需重复安装。'}
    backup = workspace / 'backups' / ('office-tool-v2-' + datetime.now().strftime('%Y%m%d-%H%M%S-%f'))
    backup.mkdir(parents=True)
    shutil.copy2(profile, backup / 'cordis.patch.yml')
    shutil.copy2(hooks_path, backup / 'hooks.json')
    if plugin.exists():
        shutil.copytree(plugin, backup / 'plugin')
    # Rename only this registered skill entry. A Windows junction itself moves;
    # its old tool/office target and the user's business folders are untouched.
    if skill.exists() or skill.is_symlink():
        if skill.parent.resolve() != (workspace / 'skills').resolve():
            raise RuntimeError('技能目录不是预期路径')
        os.rename(skill, backup / 'office-suite')
    try:
        shutil.copytree(ROOT / 'dsh-plugin', plugin, dirs_exist_ok=True)
        shutil.copytree(ROOT / 'skills/office-suite-v2', skill)
        skill_file = skill / 'SKILL.md'
        text = skill_file.read_text(encoding='utf-8')
        text = re.sub(r'(?m)^name:\s*office-suite-v2\s*$', 'name: office-suite', text)
        skill_file.write_text(text, encoding='utf-8')
        profile.write_text(wanted_profile, encoding='utf-8')
        hooks_path.write_text(json.dumps(hooks, ensure_ascii=False, indent=2), encoding='utf-8')
    except Exception as exc:
        raise RuntimeError(f'接入未完成；原配置和技能入口备份在 {backup}。原因：{exc}') from exc
    return {'ok': True, 'changed': True, 'tool_root': str(ROOT), 'backup': str(backup),
            'message': '已接入 toolV2；退出并重新打开 DSH 后生效。旧聊天和业务文件未改动。'}


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='把当前 toolV2 接入本机 DSH，不联网安装依赖')
    parser.add_argument('--dsh-root', type=Path, default=ROOT.parent / 'deepseek-harness-local')
    args = parser.parse_args()
    try:
        print(json.dumps(install(args.dsh_root), ensure_ascii=False, indent=2))
    except Exception as exc:
        print(json.dumps({'ok': False, 'error': str(exc)}, ensure_ascii=False))
        raise SystemExit(1)
