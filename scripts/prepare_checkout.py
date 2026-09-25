"""Create missing local-only DSH config without replacing existing settings."""
from pathlib import Path
import json
import subprocess
import sys

ROOT = Path(__file__).resolve().parents[1]
H = ROOT / 'deepseek-harness-local'
W = H / 'workspace'


def create(path, content):
    path.parent.mkdir(parents=True, exist_ok=True)
    if path.exists():
        return
    path.write_text(content, encoding='utf-8')


def main():
    profile = W / 'profiles/web'
    create(profile / 'cordis.yml', '[]\n')
    create(profile / 'package.json', json.dumps({
        'name': 'dsh-profile-web', 'private': True, 'dependencies': {},
        'dsh': {'profile': {'bundles': ['@deepseek-ai/dsh-base', '@deepseek-ai/dsh-web-app'],
                            'patchReload': 'live'}}}, indent=2))
    create(profile / 'pnpm-workspace.yaml', 'packages:\n  - .\nnodeLinker: hoisted\nautoInstallPeers: false\n')
    hooks = {'hooks': {event: [{'hooks': [{'type': 'command',
              'command': f'"{sys.executable}" "{ROOT / "toolV2/hooks/bank_gate.py"}"', 'timeout': 30}]}]
              for event in ('SessionStart', 'UserPromptSubmit')}}
    create(W / 'hooks.json', json.dumps(hooks, ensure_ascii=False, indent=2))
    quote = lambda p: json.dumps(str(p), ensure_ascii=False)
    create(profile / 'cordis.patch.yml',
        '- insert:\n    - id: bank-gate-hooks\n      name: "@deepseek-ai/dsh-hooks-claude-code"\n'
        '      config:\n        configPath: ' + quote(W / 'hooks.json') + '\n'
        '- insert:\n    - id: bank-approval\n      name: ' + quote(W / 'plugins/bank-approval/index.mjs') + '\n'
        '- id: ui-approval\n  disabled: true\n'
        '- insert:\n    - id: bank-approval-ui\n      name: ' + quote(W / 'plugins/bank-approval-ui/lib/index.js') + '\n')
    # Existing .env and credentials are deliberately neither read nor copied.
    subprocess.run([sys.executable, str(ROOT / 'toolV2/install.py'), '--dsh-root', str(H)], check=True)
    print('配置准备完成。已有配置未覆盖。按说明在本机配置获准模型，再启动 DSH。')


if __name__ == '__main__':
    main()
