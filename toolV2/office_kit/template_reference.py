"""Resolve an explicit reference to a fixed value in the same workbook."""
import re
from .target_validation import read_excel_template


def resolve(store, path, target, reference, label):
    if (path.suffix.lower() != '.xlsx' or target.get('kind') != 'xlsx_cell' or
            not isinstance(reference, dict) or set(reference) != {'sheet', 'cell'} or
            not all(isinstance(reference[k], str) and reference[k] for k in reference) or
            not re.fullmatch(r'[A-Za-z]{1,3}[1-9][0-9]{0,6}', reference.get('cell', ''))):
        raise ValueError('模板固定值引用需要同一 Excel 模板内的 sheet 和 cell')
    workbook = read_excel_template(path)
    try:
        cell = workbook[reference['sheet']][reference['cell']]
    except (KeyError, ValueError, TypeError) as exc:
        raise ValueError('模板固定值引用不是当前模板中的有效单元格') from exc
    text = cell.value
    match = re.fullmatch(r'\s*([^：:]+?)\s*[：:]\s*(.*?)\s*', text, re.S) if isinstance(text, str) else None
    if (not match or text.startswith('=') or not match[2] or
            re.fullmatch(r'[\s_＿—\-□☐]+', match[2])):
        raise ValueError('模板固定值引用必须是非空、非公式、带标签的固定值，不能是占位符')
    if match[1].strip() != str(label).strip():
        raise ValueError('模板固定值引用与目标位置标签不完全一致')
    source_target = {'kind': 'xlsx_cell', 'sheet': reference['sheet'], 'cell': reference['cell']}
    from .fact_catalog import target_subject_context
    source = target_subject_context(store, path, source_target)
    destination = target_subject_context(store, path, target)
    source_role = (source.get('declared_role'), source.get('declared_number'))
    target_role = (destination.get('declared_role'), destination.get('declared_number'))
    if source_role != target_role and (source_role[0] or target_role[0]):
        raise ValueError('模板固定值引用与目标位置的明确主体不同或不完整，请选择对应主体的证据')
    return {'value': match[2].strip(), 'source_kind': 'template_reference',
            'provenance': f'同模板固定值 {path.name} / {reference["sheet"]}!{reference["cell"]}'}
