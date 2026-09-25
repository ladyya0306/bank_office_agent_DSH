"""Synthetic regressions for cross-template review grouping and delivery versions."""
from __future__ import annotations

import unittest
import tempfile
from pathlib import Path
from unittest.mock import patch

from workflow import review
from workflow import report
from workflow.delivery import artifact_versions, build_delivery


def row(n, field, *, fact_id, purpose=None, value='合成值', entity=7,
        subject='borrower', risk=None, local_issue=None):
    return {'n': n, 'field': field, 'value': value, 'entity_name': '合成主体',
            'entity_id': entity, 'subject_eid': entity, 'subject_scope': subject,
            'provenance': '合成材料第1行', 'source_kind': 'source',
            'ambiguous': False, 'candidates': [], 'high_risk': risk,
            'local_issue': local_issue,
            'target': {'kind': 'anchor', 'anchor': f'填充位置{n}'},
            '_test_meta': {'fact_id': fact_id, '_base_field': purpose or field}}


class ReviewGroupingTests(unittest.TestCase):
    def question_set(self, rows):
        docs = {}
        for i, item in enumerate(rows):
            name = f'表{i}.docx'
            docs[name] = {'status': 'pending', 'plan': {'rows': [item]}}
        task = {'batch': '20260925-01', 'documents': docs}
        with patch.object(review, '_fact_catalog', return_value={
                item['field']: item['_test_meta'] for item in rows}), \
             patch.object(review, 'saved_choices', side_effect=lambda p: ({}, [p['rows'][0]['n']])), \
             patch.object(review, 'fingerprint', side_effect=lambda p, r: str(r['n'])):
            return review.questions(task)

    def test_aliases_of_same_fact_and_same_purpose_merge_with_detail(self):
        questions = self.question_set([row(1, '借款人联系电话', fact_id=10, purpose='联系电话'),
                                       row(2, '联系电话', fact_id=10)])
        self.assertEqual(1, len(questions))
        self.assertEqual(2, len(questions[0]['locations']))
        self.assertIn('建议填写', questions[0]['question'])
        self.assertIn('填写位置（2 处）', questions[0]['detail'])

    def test_same_value_does_not_merge_across_fact_purpose_or_owner(self):
        variants = [row(1, '联系电话', fact_id=10),
                    row(2, '开户行联系电话', fact_id=10, purpose='开户行联系电话'),
                    row(3, '电话', fact_id=11, purpose='联系电话'),
                    row(4, '联系电话', fact_id=10, entity=8)]
        self.assertEqual(4, len(self.question_set(variants)))

    def test_changed_value_is_asked_again(self):
        first = self.question_set([row(1, '联系电话', fact_id=10, value='13800000000')])
        changed = self.question_set([row(1, '联系电话', fact_id=10, value='13900000000')])
        self.assertEqual(first[0]['id'], changed[0]['id'])
        self.assertIn('13900000000', changed[0]['question'])

    def test_local_position_issue_stays_separate_and_only_suggests_blank(self):
        questions = self.question_set([row(1, '联系电话', fact_id=10),
                                       row(2, '联系电话', fact_id=10,
                                           local_issue='此处是日期格式，建议留空')])
        issue = next(q for q in questions if '日期格式' in q['question'])
        self.assertEqual(['留空'], [o['label'] for o in issue['options']])
        self.assertEqual(1, len(issue['locations']))

    def test_unresolved_owner_without_identity_is_isolated_by_template(self):
        items = [row(1, '联系电话', fact_id=None, entity=None),
                 row(2, '联系电话', fact_id=None, entity=None)]
        self.assertEqual(2, len(self.question_set(items)))

    def test_resolved_subject_uses_exact_fact_from_ambiguous_global_catalog(self):
        raw = row(1, '联系电话', fact_id=None)
        raw['_test_meta']['_candidates'] = [
            {'fact_id': 10, 'entity_id': 7, 'value': '合成值', 'provenance': '合成材料第1行'},
            {'fact_id': 11, 'entity_id': 8, 'value': '合成值', 'provenance': '合成材料第2行'}]
        alias = row(2, '借款人联系电话', fact_id=10, purpose='联系电话')
        self.assertEqual(1, len(self.question_set([raw, alias])))
        raw['_test_meta']['_candidates'][0]['provenance'] = '不同来源'
        self.assertEqual(2, len(self.question_set([raw, alias])))


class DeliveryTests(unittest.TestCase):
    def setUp(self):
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        self.root = Path(tmp.name).resolve()
        self.work = str(self.root / 'work')

    def test_current_and_superseded_versions_and_eight_item_groups(self):
        records = []
        docs = {}
        for i in range(9):
            name = f'templates/模板{i}.docx'
            current = f'out/current-{i}.docx'
            docs[name] = {'status': 'completed', 'output': current}
            records.append({'template_path': f'{self.work}/{name}', 'artifact_path': f'{self.work}/{current}',
                            'artifact_sha256': f'new-{i}', 'run_id': f'run-new-{i}'})
            records.append({'template_path': f'{self.work}/{name}', 'artifact_path': f'{self.work}/out/old-{i}.docx',
                            'artifact_sha256': f'old-{i}', 'run_id': f'run-old-{i}'})
        task = {'last_work': self.work, 'documents': docs}
        result = build_delivery(task, records, f'{self.work}/report.md')
        self.assertEqual(9, sum(v['status'] == 'current' for v in result['versions']))
        self.assertEqual(9, sum(v['status'] == 'superseded' for v in result['versions']))
        self.assertEqual([8, 2], [len(g) for g in result['attachment_groups']])
        self.assertEqual(10, sum(map(len, result['attachment_groups'])))
        self.assertTrue(all(len(g) <= 8 for g in result['attachment_groups']))
        self.assertTrue(all(Path(path).is_relative_to(self.work)
                            for path in result['attachments']))

    def test_same_template_name_in_different_directories_does_not_match(self):
        task = {'last_work': self.work, 'documents': {
            'forms/current.docx': {'status': 'completed', 'output': 'out/current.docx'},
            'archive/current.docx': {'status': 'completed', 'output': 'out/archive.docx'},
        }}
        records = [{'template_path': f'{self.work}/archive/current.docx',
                    'artifact_path': f'{self.work}/out/archive.docx', 'run_id': 'right'},
                   {'template_path': str(self.root / 'else/archive/current.docx'),
                    'artifact_path': str(self.root / 'else/out/archive.docx'), 'run_id': 'other-task'}]
        versions = artifact_versions(task, records)
        self.assertEqual('current', next(v['status'] for v in versions
                                         if v['template_path'] == 'archive/current.docx'))
        self.assertEqual('historical', next(v['status'] for v in versions if v['run_id'] == 'other-task'))

    def test_history_and_registered_artifact_deduplicate_and_previous_is_not_delivered(self):
        task = {'last_work': self.work, 'documents': {
            'form.docx': {'status': 'pending', 'output_history': [
                {'output': 'out/old.docx', 'output_hash': 'old-hash', 'run_id': 'old-run'}]},
        }}
        records = [{'template_path': f'{self.work}/form.docx', 'artifact_path': f'{self.work}/out/old.docx',
                    'artifact_sha256': 'old-hash', 'run_id': 'old-run'}]
        result = build_delivery(task, records, f'{self.work}/report.md', max_attachments=20)
        self.assertEqual(1, len(result['versions']))
        self.assertEqual('previous', result['versions'][0]['status'])
        self.assertEqual([str(Path(self.work) / 'report.md')], result['attachments'])
        self.assertTrue(all(len(group) <= 8 for group in result['attachment_groups']))

    def test_report_explains_user_blank_and_only_prints_recorded_stage_times(self):
        with tempfile.TemporaryDirectory() as tmp:
            target = {'kind': 'anchor', 'anchor': '期限年数'}
            row_data = {'kind': 'slot', 'n': 1, 'field': '期限年数',
                        'label': '期限年数', 'target': target, 'value': '2026至2028',
                        'local_issue': '该位置需要年数，来源只有日期区间'}
            task = {'id': 'synthetic', 'batch': '20260925-01', 'status': 'completed',
                    'documents': {'期限表.docx': {'status': 'completed', 'filled': 0,
                                                   'plan': {'db_path': 'synthetic', 'rows': [row_data]}}},
                    'timing': {'stages': {'source_parse_and_check': {'seconds': 2.5, 'count': 1},
                                          'user_confirmation_wait': {'seconds': 8.0, 'count': 1}}}}
            with patch.object(report, 'saved_choices', return_value=({'blank': ['1']}, [])):
                report.write(Path(tmp), task)
            text = (Path(tmp) / task['report']).read_text(encoding='utf-8')
            self.assertIn('用户确认留空', text)
            self.assertIn('来源只有日期区间', text)
            self.assertNotIn('来源未给出建议值', text)
            self.assertIn('2.5 秒', text)
            self.assertIn('8.0 秒', text)
            self.assertNotIn('位置整理与检查：0 秒', text)


if __name__ == '__main__':
    unittest.main()
