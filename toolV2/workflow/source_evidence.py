"""Validate dynamic-script source claims against task-owned Office material.

This module only reads the source snapshot that the task registered.  It does
not decide ownership, write facts, or change task state; callers can therefore
validate an entire proposed batch before choosing whether to persist it.
"""
from __future__ import annotations

import hmac
import re
from pathlib import Path
from typing import Any

from office_kit.store_v2 import sha256_file

from .storage import inside


def _work_path(work: str | Path) -> Path:
    return Path(work).resolve(strict=True)


def _source_path(work: Path, task: dict[str, Any], source: str) -> Path:
    if not isinstance(source, str) or source not in task.get("source", []):
        raise ValueError("更新指定的来源不属于当前任务：%r" % source)
    return inside(work, source)


def _normal(text: Any) -> str:
    return re.sub(r"\s+", "", str(text))


def _nonempty_text(value: Any, name: str) -> str:
    if value is None or not str(value).strip():
        raise ValueError("%s 不能为空" % name)
    return str(value)


def _docx_items(path: Path) -> list[dict[str, Any]]:
    from office_kit.target_validation import read_word_template

    engine = read_word_template(path)
    indexes: dict[str, int] = {}
    items: list[dict[str, Any]] = []
    for element, paragraph in engine._entries:
        part = engine._part_by_el.get(element,
                                      engine._part_by_el.get(paragraph._element, "word/document.xml"))
        index = indexes.get(part, 0)
        indexes[part] = index + 1
        raw = engine.xml_for(paragraph)
        text = raw.text if raw is not None else paragraph.text
        if text is not None and str(text).strip():
            items.append({"part": part, "paragraph_index": index, "text": str(text)})
    return items


def document_items(work: str | Path, task: dict[str, Any], source: str) -> list[dict[str, Any]]:
    """Return nonempty original source items for one task-registered file.

    Returned items retain the physical Office position and text.  This is a
    read-only evidence view, rather than a field extraction or ownership rule.
    """
    root = _work_path(work)
    path = _source_path(root, task, source)
    suffix = path.suffix.lower()
    if suffix in (".xlsx", ".xlsm"):
        import openpyxl

        workbook = openpyxl.load_workbook(str(path), read_only=False, data_only=False)
        try:
            items = []
            for sheet in workbook.worksheets:
                for row in sheet.iter_rows():
                    for cell in row:
                        if cell.value is not None and str(cell.value).strip():
                            items.append({"sheet": sheet.title, "cell": cell.coordinate,
                                          "text": str(cell.value)})
            return items
        finally:
            workbook.close()
    if suffix == ".docx":
        return _docx_items(path)
    raise ValueError("动态脚本证据仅支持 DOCX、XLSX 或 XLSM 来源：%s" % path.name)


def _evidence_item(path: Path, evidence: Any, items: list[dict]) -> dict:
    if not isinstance(evidence, dict) or not evidence:
        raise ValueError("每条来源行必须提供证据位置")
    if path.suffix.lower() in (".xlsx", ".xlsm"):
        if set(evidence) != {"sheet", "cell"} or not all(isinstance(v, str) and v for v in evidence.values()):
            raise ValueError("XLSX 证据需要 sheet 和 cell")
        match = next((item for item in items if item['sheet'] == evidence['sheet']
                      and item['cell'] == evidence['cell'].upper()), None)
    else:
        index = evidence.get('paragraph_index')
        if (set(evidence) - {'paragraph_index', 'part'} or isinstance(index, bool)
                or not isinstance(index, int) or index < 0):
            raise ValueError('DOCX 证据需要非负 paragraph_index 和可选 part')
        match = next((item for item in items if item['part'] == evidence.get('part', 'word/document.xml')
                      and item['paragraph_index'] == index), None)
    if match is None:
        raise ValueError('证据位置不存在或内容为空：%s' % evidence)
    return match


def validate_updates(work: str | Path, task: dict[str, Any], updates: Any) -> list[dict[str, Any]]:
    """Validate dynamic-script source claims and return standard absorb rows.

    The function is intentionally all-or-nothing: it only returns after every
    source hash, field, evidence location, literal value, and supplied entity
    has been checked.  It does not mutate ``task`` or persistent storage.
    """
    if not isinstance(updates, list) or not updates:
        raise ValueError("来源更新必须是非空列表")
    root = _work_path(work)
    output: list[dict[str, Any]] = []
    source_texts: dict[str, str] = {}

    for update_index, update in enumerate(updates, start=1):
        if not isinstance(update, dict):
            raise ValueError("第%d个来源更新必须是对象" % update_index)
        source = update.get("source")
        if source in source_texts:
            raise ValueError("同一来源在一次提交中重复出现；请合并 rows：%s" % source)
        path = _source_path(root, task, source)
        supplied_hash = update.get("source_sha256")
        if not isinstance(supplied_hash, str) or not supplied_hash:
            raise ValueError("来源 %s 缺少 source_sha256" % source)
        actual_hash = sha256_file(path)
        if not hmac.compare_digest(supplied_hash.lower(), actual_hash.lower()):
            raise ValueError("来源 %s 的 source_sha256 与当前文件不一致；未采用任何更新" % source)
        rows = update.get("rows")
        if not isinstance(rows, list) or not rows:
            raise ValueError("来源 %s 的 rows 必须是非空列表" % source)
        if source not in source_texts:
            source_items = document_items(root, task, source)
            source_texts[source] = "\n".join(item["text"] for item in source_items)

        for row_index, proposed in enumerate(rows, start=1):
            if not isinstance(proposed, dict):
                raise ValueError("来源 %s 的第%d条 rows 必须是对象" % (source, row_index))
            key = proposed.get("key")
            if not isinstance(key, str) or not key.strip():
                raise ValueError("来源 %s 的第%d条 rows 缺少非空 key" % (source, row_index))
            value = _nonempty_text(proposed.get("value"), "来源 %s 的第%d条 value" % (source, row_index))
            item = _evidence_item(path, proposed.get("evidence"), source_items)
            derivation = proposed.get('derivation')
            if derivation is not None and (not isinstance(derivation, str) or not derivation.strip()):
                raise ValueError('转换或计算说明 derivation 必须为非空文字')
            if _normal(value) not in _normal(item["text"]) and not derivation:
                raise ValueError("来源 %s 的第%d条 value 不是证据原文的字面子串" %
                                 (source, row_index))
            entity_name = proposed.get("entity_name")
            if "entity_name" in proposed:
                if not isinstance(entity_name, str) or not entity_name.strip():
                    raise ValueError("来源 %s 的第%d条 entity_name 必须为非空字符串" %
                                     (source, row_index))
                if _normal(entity_name) not in _normal(source_texts[source]):
                    raise ValueError("来源 %s 的第%d条 entity_name 未在该来源原文中出现" %
                                     (source, row_index))
            role = proposed.get("role")
            if role is not None and (not isinstance(role, str) or not role.strip()):
                raise ValueError("来源 %s 的第%d条 role 必须为非空字符串" % (source, row_index))
            evidence = dict(proposed["evidence"])
            line = ("%s!%s" % (item["sheet"], item["cell"])
                    if "sheet" in item else "%s#%d" % (item["part"], item["paragraph_index"]))
            output.append({
                "n": len(output) + 1,
                "line": line,
                "label": key,
                "key": key,
                "value": value,
                "quote": item["text"],
                "_source": source,
                "source_sha256": actual_hash,
                "evidence": evidence,
                "derivation": derivation,
                "entity_name": entity_name if "entity_name" in proposed else None,
                "role": role,
                "kind": "字段",
                "owner_kind": "subject" if entity_name else "unknown",
                "context": "动态脚本声明的来源位置已由宿主核验",
                "confidence": 1.0,
                "needs": "收",
                "assumed": False,
                "known_key": True,
                "ownership_group": None,
                "certificate_type": None,
                "owner_heading": None,
            })
    return output
