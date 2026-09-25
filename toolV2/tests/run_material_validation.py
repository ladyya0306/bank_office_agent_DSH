"""Offline maintainer validation. Only synthetic fixtures; failures remain failures."""
import datetime
import hashlib
import importlib.metadata
import json
from pathlib import Path
import platform
import subprocess
import sys

ROOT = Path(__file__).resolve().parents[1]


def snapshot():
    files = [ROOT / 'office.py', ROOT / 'install.py']
    for folder in ('workflow', 'office_kit', 'dsh-plugin', 'skills', 'tests'):
        files.extend(p for p in (ROOT / folder).rglob('*')
                     if p.is_file() and p.suffix in ('.py', '.mjs', '.md'))
    hashes = {p.relative_to(ROOT).as_posix(): hashlib.sha256(p.read_bytes()).hexdigest()
              for p in sorted(files)}
    return {'time': datetime.datetime.now().astimezone().isoformat(),
            'python': platform.python_version(), 'platform': platform.platform(),
            'dependencies': {name: importlib.metadata.version(name)
                             for name in ('python-docx', 'openpyxl')},
            'code_and_tests_sha256': hashes}


def main():
    evidence = ROOT / 'validation_runs' / datetime.datetime.now().strftime('%Y%m%d-%H%M%S-%f')
    evidence.mkdir(parents=True)
    commands = [[sys.executable, 'tests/' + name] for name in (
        'test_material_variants.py', 'test_source_roles.py', 'validate_target_variants.py',
        'test_workflow.py', 'test_source_duplicates.py', 'acceptance_batch.py',
        'test_source_repair.py', 'test_source_diagnostics_workflow.py',
        'test_header_filling.py', 'test_target_labels.py', 'test_multiple_fill_locations.py',
        'test_fill_review_recovery.py')]
    commands += [[sys.executable, '-m', 'pytest', 'tests/test_template_subject.py', '-q']]
    commands += [['node', '--test', 'tests/plugin-controller.mjs'],
                 ['node', '--test', 'tests/plugin-workspaces.mjs']]
    report = snapshot()
    report['runs'] = []
    for i, command in enumerate(commands):
        try:
            cp = subprocess.run(command, cwd=ROOT, capture_output=True,
                                encoding='utf-8', errors='replace', timeout=300)
            output, code = cp.stdout + '\n' + cp.stderr, cp.returncode
        except (OSError, subprocess.TimeoutExpired) as exc:
            output, code = str(exc), -1
        logfile = evidence / f'{i + 1:02d}.txt'
        logfile.write_text(output, encoding='utf-8')
        report['runs'].append({'command': command, 'exit_code': code, 'log': logfile.name})
        print(f'{command[-1]}: {"PASS" if code == 0 else "FAIL"}', flush=True)
    report['passed'] = all(run['exit_code'] == 0 for run in report['runs'])
    (evidence / 'report.json').write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')
    print(str(evidence))
    return 0 if report['passed'] else 1


if __name__ == '__main__':
    sys.exit(main())
