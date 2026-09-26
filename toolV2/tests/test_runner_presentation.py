"""Unmapped positions are navigation data, not duplicated error fragments."""
from __future__ import annotations

import tempfile
from pathlib import Path

from docx import Document

from office_kit.store_v2 import StoreV2
from workflow.runner import dispatch


def make_doc(path: Path, *lines: str) -> None:
    doc = Document()
    for line in lines:
        doc.add_paragraph(line)
    doc.save(path)


def test_unmapped_slots_are_not_issues_but_invalid_saved_rule_is():
    with tempfile.TemporaryDirectory() as raw:
        work = Path(raw)
        make_doc(work / 'source.docx', '借款人：合成公司', '联系电话：13800000000')
        make_doc(work / 'unmapped.docx', '未知业务项目：____')
        result = dispatch({'action': 'start', 'work': str(work), 'source': ['source.docx'],
                           'targets': ['unmapped.docx'], 'batch': '20260926-99'})
        assert result['status'] == 'needs_mapping'
        assert result['issues'] == []
        assert result['mapping_requests'][0]['positions']

        repeat = dispatch({'action': 'status', 'work': str(work), 'task_id': result['task_id']})
        assert repeat['status'] == 'needs_mapping'
        assert repeat['issues'] == []
        assert repeat['counters'] == result['counters']

        # Inject a legacy invalid target to verify real validation errors stay visible.
        with StoreV2(work / 'db/workflow.db', actor='synthetic-test') as store:
            tid = store.register_template(work / 'unmapped.docx', result['batch'])
            store.add_rule(tid, '联系电话', '联系电话',
                           {'kind': 'anchor', 'anchor': '不存在的锚点：', 'max_blank': 10},
                           confidence=1, decided_by='human', batch_no=result['batch'])
        invalid = dispatch({'action': 'status', 'work': str(work), 'task_id': result['task_id']})
        assert invalid['status'] == 'needs_mapping'
        assert invalid['issues']
