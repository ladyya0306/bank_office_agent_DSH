"""Pure helpers for artifact version labels and DSH-sized attachment groups."""
from __future__ import annotations

from pathlib import Path, PurePath
import os


MAX_ATTACHMENTS = 8


def _same_path(left, right):
    if not left or not right:
        return False
    left = str(left).replace('\\', '/').rstrip('/').casefold()
    right = str(right).replace('\\', '/').rstrip('/').casefold()
    return left == right or left.endswith('/' + right.lstrip('/')) or right.endswith('/' + left.lstrip('/'))


def _absolute(task, path):
    path = Path(path)
    if not path.is_absolute() and task.get('last_work'):
        path = Path(task['last_work']) / path
    return os.path.abspath(str(path))


def _path_key(task, path):
    return os.path.normcase(os.path.normpath(_absolute(task, path)))


def _template_key(task, template_path):
    if not template_path:
        return None
    candidate = Path(template_path)
    root = Path(task['last_work']) if task.get('last_work') else None
    if root and not candidate.is_absolute():
        candidate = root / candidate
    try:
        if root:
            return os.path.normcase(os.path.normpath(candidate.resolve(strict=False).relative_to(
                root.resolve(strict=False)).as_posix()))
    except ValueError:
        pass
    return os.path.normcase(os.path.normpath(str(candidate.resolve(strict=False))))


def artifact_versions(task, artifact_records):
    """Return registered files labelled current or superseded per template."""
    current = {os.path.normcase(os.path.normpath(name)): rec.get('output')
               for name, rec in task.get('documents', {}).items()
               if rec.get('output') and rec.get('status') == 'completed'}
    output_by_template = {os.path.normcase(os.path.normpath(name)): rec
                          for name, rec in task.get('documents', {}).items()}
    result = []
    seen = set()
    for name, record in task.get('documents', {}).items():
        template_key = os.path.normcase(os.path.normpath(name))
        has_current = template_key in current
        for old in record.get('output_history', []):
            path = old.get('output') or old.get('path')
            if path:
                if has_current and _path_key(task, path) == _path_key(task, current[template_key]):
                    continue
                item = {'template': name, 'template_path': name,
                        'path': _absolute(task, path),
                        'sha256': old.get('output_hash') or old.get('sha256'),
                        'run_id': old.get('run_id'),
                        'status': 'superseded' if has_current else 'previous'}
                result.append(item)
                seen.add((template_key, _path_key(task, path)))
        if template_key in current:
            path = current[template_key]
            item = {'template': name, 'template_path': name, 'path': _absolute(task, path),
                    'sha256': record.get('output_hash'), 'run_id': record.get('run_id'),
                    'status': 'current'}
            result.append(item)
            seen.add((template_key, _path_key(task, path)))
    for artifact in artifact_records or []:
        path = artifact.get('artifact_path')
        template_key = _template_key(task, artifact.get('template_path'))
        template = artifact.get('template_path')
        if not path:
            continue
        if template_key is None or template_key not in output_by_template:
            status = 'historical'
        else:
            rec = output_by_template[template_key]
            if rec.get('status') != 'completed' or not rec.get('output'):
                status = 'previous'
            else:
                status = 'current' if _path_key(task, path) == _path_key(task, rec['output']) else 'superseded'
        identity = (template_key, _path_key(task, path))
        if identity in seen:
            continue
        result.append({'template': template or PurePath(path).name,
                       'template_path': template, 'path': _absolute(task, path),
                       'sha256': artifact.get('artifact_sha256'),
                       'run_id': artifact.get('run_id'), 'status': status})
        seen.add(identity)
    return result


def attachment_groups(attachments, *, max_attachments=8):
    """Split attachment paths into ordered groups, counting report attachments too."""
    if not isinstance(max_attachments, int) or max_attachments < 1:
        raise ValueError('max_attachments 必须是正整数')
    max_attachments = min(max_attachments, MAX_ATTACHMENTS)
    items = list(attachments or [])
    return [items[i:i + max_attachments] for i in range(0, len(items), max_attachments)]


def build_delivery(task, artifact_records, report_path=None, *, max_attachments=8):
    """Create response-ready delivery metadata without changing files or task state."""
    if not isinstance(max_attachments, int) or max_attachments < 1:
        raise ValueError('max_attachments 必须是正整数')
    effective_limit = min(max_attachments, MAX_ATTACHMENTS)
    versions = artifact_versions(task, artifact_records)
    current = [v for v in versions if v['status'] == 'current']
    paths = [v['path'] for v in current]
    if report_path:
        paths.append(_absolute(task, report_path))
    return {'versions': versions, 'attachments': paths,
            'attachment_groups': attachment_groups(paths, max_attachments=effective_limit),
            'max_attachments_per_group': effective_limit}
