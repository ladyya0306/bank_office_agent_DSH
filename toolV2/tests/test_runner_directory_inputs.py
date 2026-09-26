"""Directory arguments normalize to the same safe task inputs as file lists."""
from __future__ import annotations

import os
from pathlib import Path

import pytest
from docx import Document

from workflow.runner import dispatch


def write_doc(path: Path, *lines: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    doc = Document()
    for line in lines:
        doc.add_paragraph(line)
    doc.save(path)


def answer_all(result: dict) -> dict:
    while result['status'] in ('awaiting_source', 'awaiting_fill'):
        result = dispatch({'action': 'resume', 'work': result['work'], 'task_id': result['task_id'],
                           'answers': [{'id': q['id'], 'selected': [q['options'][0]['label']], 'custom': ''}
                                       for q in result['questions']]})
    return result


def test_directory_inputs_complete_and_reuse_explicit_file_task(tmp_path: Path):
    source_dir = tmp_path / 'materials' / 'nested'
    target_dir = tmp_path / 'templates' / 'nested'
    source_file = source_dir / 'source.docx'
    target_file = target_dir / 'template.docx'
    write_doc(source_file, '\u501f\u6b3e\u4eba\u540d\u79f0\uff1a\u5408\u6210\u76ee\u5f55\u516c\u53f8')
    write_doc(target_file, '\u501f\u6b3e\u4eba\u540d\u79f0\uff1a____')
    write_doc(target_dir / '~$template.docx', 'Office lock file must be ignored')
    (target_dir / 'notes.txt').write_text('not a template', encoding='utf-8')

    directory = answer_all(dispatch({'action': 'start', 'work': str(tmp_path),
                                     'source': ['materials/nested'], 'targets': ['templates/nested'],
                                     'batch': '20260926-98'}))
    assert directory['status'] == 'completed', directory.get('error')
    assert directory['counters']['source_reads'] == 1

    explicit = answer_all(dispatch({'action': 'start', 'work': str(tmp_path),
                                    'source': ['materials/nested/source.docx'],
                                    'targets': ['templates/nested/template.docx']}))
    assert explicit['status'] == 'completed'
    assert explicit['task_id'] == directory['task_id']
    assert explicit['batch'] == directory['batch']
    assert explicit['counters'] == directory['counters']


def test_directory_outside_workspace_and_symlink_are_rejected(tmp_path: Path):
    outside = tmp_path.parent / f'{tmp_path.name}-outside'
    outside.mkdir()
    write_doc(outside / 'source.docx', '借款人名称：工作区外')
    write_doc(tmp_path / 'target.docx', '无待填位置')
    payload = {'action': 'start', 'work': str(tmp_path), 'targets': ['target.docx']}
    with pytest.raises(ValueError, match='不在本工作区'):
        dispatch({**payload, 'source': [str(outside)]})

    link = tmp_path / 'linked-outside'
    try:
        os.symlink(outside, link, target_is_directory=True)
    except OSError as exc:
        pytest.skip(f'current Windows test environment cannot create directory symlinks: {exc}')
    with pytest.raises(ValueError, match='不在本工作区'):
        dispatch({**payload, 'source': ['linked-outside']})
