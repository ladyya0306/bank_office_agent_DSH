from docx import Document

from office_kit.store_v2 import StoreV2
from workflow.mapping import inspect
from workflow.mapping_view import page


def test_numbered_guarantor_subject_is_visible_across_mapping_pages(tmp_path):
    path = tmp_path / 'synthetic-guarantees.docx'
    doc = Document()
    for index in range(24):
        number = 1 if index % 2 == 0 else 2
        doc.add_paragraph(f'保证人{number}名称 ____ 位置{index:02d}')
    doc.save(path)

    with StoreV2(tmp_path / 'synthetic.db') as store:
        store.batch_scope = 'synthetic-numbered-subjects'
        for number, name in ((1, '合成保证人甲公司'), (2, '合成保证人乙公司')):
            eid, _ = store.ensure_entity(name)
            store.set_role(eid, '保证人', evidence=f'保证人{number}',
                           batch_no=store.batch_scope)

        record = {'status': 'needs_mapping'}
        missing = inspect(store, path, record, {})
        assert len(missing) == 24
        assert {slot['target_subject']['number'] for slot in missing} == {1, 2}
        assert {tuple(slot['target_subject']['entity_names']) for slot in missing} == {
            ('合成保证人甲公司',), ('合成保证人乙公司',)
        }

        task = {'documents': {path.name: record}, 'available_fields': []}
        query = {'section': 'positions', 'template': path.name}
        visible = []
        page_count = 0
        while query:
            result = page(task, query)
            visible.extend(result['items'])
            query = result['next']
            page_count += 1

        assert page_count > 1
        assert len(visible) == 24
        assert all(item['target_subject']['role'] == '保证人' for item in visible)
        assert {item['target_subject']['number'] for item in visible} == {1, 2}
        for item in visible:
            expected = ('合成保证人甲公司' if item['target_subject']['number'] == 1
                        else '合成保证人乙公司')
            assert item['target_subject']['entity_names'] == [expected]
