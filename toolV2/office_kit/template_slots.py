"""Read-only discovery of addressable blank locations in Office templates.

This module intentionally discovers locations only.  It does not infer facts,
does not write Office packages, and does not treat a reference answer as data.
"""
from __future__ import annotations

import hashlib
from html import unescape as html_unescape
import json
import re
import zipfile
from pathlib import Path
from typing import Any

from .xml_fill import scan_paragraphs

_PART_RE = re.compile(r"word/(?:document|header\d+|footer\d+)\.xml$")
_UNDERLINE_RE = re.compile(r"_{2,}")
_BRACKET_RE = re.compile(r"(?<=[(（])[ _\u3000]*(?=[)）])")
_COLON_END_RE = re.compile(r"[：:]([ \u3000]*)$")
_LABEL_RE = re.compile(r"(?:^|[\n，、；;。])\s*([^\n，、；;。:：]{1,40})[：:]")
_PROTECTED_RE = re.compile(r"签字|签名|签章|盖章|签核|审批|审核|经办人|法定代表人.{0,8}(?:签|章)")
_DATE_LINE_RE = re.compile(r"^\s*(?:\d{4}|年\s*月\s*日|日期)\s*[年月日\-/. ]*\s*$")
_SECTION_LABEL_RE = re.compile(r"(?:填写|说明|意见|核准|审批|调查|备注|决议|会议|日期)$")
_SUFFIX_RE = re.compile(r"(?P<slot>[ \u3000]+)(?=(?:同志|公司|有限公司|有限责任公司|借款人|万元|元|人|名|年|个月|董事会))")
_PLACEHOLDER_PREFIX_RE = re.compile(r"(?:经核实[，,]|编号为|会议名称[：:]|借款人[：:]|保证人[：:]|人民币|金额|期限|人数|年限)\s*$")
_RUN_RE = re.compile(rb"<w:r(?:\s[^>]*)?>(.*?)</w:r>", re.S)
_TEXT_RE = re.compile(rb"<w:t(?:\s[^>]*)?>(.*?)</w:t>", re.S)


def _bounded(value: str, limit: int = 600) -> str:
    return value if len(value) <= limit else value[:limit] + "…（原段过长，已截断）"


def _label(text: str, start: int) -> str:
    left = text[:start]
    match = list(_LABEL_RE.finditer(left))
    if match:
        return match[-1].group(1).strip().lstrip("*＊")
    tail = re.split(r"[\n，、；;。]", left)[-1].strip(" ：:\u3000")
    if tail:
        return tail[-40:]
    # A line may deliberately begin with a blank, as in "____同志" or
    # "____公司".  Preserve that useful signal rather than classifying it as
    # pure indentation.
    right = text[start:]
    suffix = re.match(r"[_\s\u3000]*(同志|公司|借款人)", right)
    return (suffix.group(1) + "（句首空位）") if suffix else "待确认空位"


def _protected(text: str, label: str, start: int = 0, end: int = 0) -> tuple[bool, str | None]:
    # Protection is local to the blank.  A statement such as "法定代表人签字
    # 真实有效" elsewhere in a paragraph must not protect an earlier contract
    # number or company-name field.
    before = text[max(0, start - 16):start] if text else ""
    after = text[end:min(len(text), end + 12)] if text else ""
    # The protected token has to be the label itself, immediately before the
    # blank, or the first thing after it.  Do not search the whole paragraph.
    immediate = bool(re.search(r"(?:签字|签名|签章|盖章|签核|审批|审核|经办人)\s*[：:]?\s*$", before))
    immediate = immediate or bool(re.match(r"\s*(?:签字|签名|签章|盖章|签核|审批|审核|公章)", after))
    probe = label if text else label
    if not (_PROTECTED_RE.search(probe) or immediate):
        return False, None
    if "公章" in probe or "单位" in probe and "章" in probe:
        return True, "单位公章位置：公司名称不能代替盖章或签章"
    return True, "签字、签核、审批或盖章位置：默认不得自动填写"


def _slot(part: str, paragraph_index: int, raw: str, start: int, end: int,
          context: dict[str, str], label: str | None = None, protected: bool | None = None,
          protected_reason: str | None = None) -> dict[str, Any]:
    label = label or _label(raw, start)
    if protected is None:
        protected, protected_reason = _protected(raw, label)
    target = {"kind": "anchor", "part": part, "paragraph_index": paragraph_index,
              "expected_text": raw, "span_start": start, "span_end": end}
    identity = json.dumps(target, ensure_ascii=False, sort_keys=True, separators=(",", ":"))
    return {"id": hashlib.sha256(identity.encode("utf-8")).hexdigest()[:20], "label": label,
            "context": context, "target": target, "protected": protected,
            "protected_reason": protected_reason}


def _underlined_space_spans(paragraph: Any, xml: bytes) -> set[tuple[int, int]]:
    """Return spaces belonging to an underlined run, not all spaces in its paragraph."""
    result: set[tuple[int, int]] = set()
    raw_xml = xml[paragraph.start:paragraph.end]
    cursor = 0
    for run in _RUN_RE.finditer(raw_xml):
        run_text = "".join(html_unescape(item.decode("utf-8", "replace")) for item in _TEXT_RE.findall(run.group(1)))
        if not run_text:
            continue
        pos = paragraph.text.find(run_text, cursor)
        if pos < 0:
            continue
        cursor = pos + len(run_text)
        if b"<w:u" not in run.group(1):
            continue
        for match in re.finditer(r" +", run_text):
            result.add((pos + match.start(), pos + match.end()))
    return result


def _normalise_spans(spans: set[tuple[int, int]]) -> list[tuple[int, int]]:
    """Coalesce one physical blank and discard its redundant colon insertion."""
    ranges = sorted((start, end) for start, end in spans if end > start)
    merged: list[tuple[int, int]] = []
    for start, end in ranges:
        if merged and start <= merged[-1][1]:
            merged[-1] = (merged[-1][0], max(merged[-1][1], end))
        else:
            merged.append((start, end))
    points = sorted((start, end) for start, end in spans if start == end)
    return merged + [point for point in points if not any(start <= point[0] <= end for start, end in merged)]


def _short_table_label(value: str) -> str | None:
    """Accept a compact one/two-paragraph table label, reject explanatory prose."""
    lines = [line.strip() for line in value.splitlines() if line.strip()]
    compact = "".join(lines).rstrip("：:").strip()
    if not compact or len(lines) > 2 or len(compact) > 40:
        return None
    if any(len(line) > 24 for line in lines) or re.search(r"[。；;，,]", compact):
        return None
    return compact


def _paragraph_slots(part: str, paragraphs: list[Any], xml: bytes) -> list[dict[str, Any]]:
    found: list[dict[str, Any]] = []
    for index, paragraph in enumerate(paragraphs):
        raw = paragraph.text
        # Whitespace-only paragraphs are normal indentation or deliberately empty
        # layout, and cannot be safely distinguished from a visual spacer.
        if not raw.strip() or _DATE_LINE_RE.match(raw):
            continue
        context = {"previous": _bounded(paragraphs[index - 1].text) if index else "",
                   "paragraph": _bounded(raw),
                   "next": _bounded(paragraphs[index + 1].text) if index + 1 < len(paragraphs) else ""}
        spans: set[tuple[int, int]] = set()
        spans.update(m.span() for m in _UNDERLINE_RE.finditer(raw))
        # A single space is normally typography.  It becomes a fillable blank
        # only when this exact paragraph carries underline formatting.
        spans.update(_underlined_space_spans(paragraph, xml))
        for match in _SUFFIX_RE.finditer(raw):
            if _PLACEHOLDER_PREFIX_RE.search(raw[max(0, match.start() - 28):match.start()]):
                spans.add(match.span("slot"))
        spans.update(m.span() for m in _BRACKET_RE.finditer(raw))
        # A label followed immediately by a colon has no characters to replace;
        # an insertion span is still precise and is handled by the XML writer.
        for match in _COLON_END_RE.finditer(raw):
            label = _label(raw, match.start())
            if label and not _SECTION_LABEL_RE.search(label):
                spans.add((match.end(), match.end()))
        for start, end in _normalise_spans(spans):
            if start == end and not raw[:start].strip():
                continue
            label = _label(raw, start)
            protected, reason = _protected(raw, label, start, end)
            found.append(_slot(part, index, raw, start, end, context, label, protected, reason))
    return found


def _word_cell_slots(path: Path) -> list[dict[str, Any]]:
    """Find empty Word cells only when a left-hand label makes intent explicit."""
    from docx import Document

    document = Document(str(path))
    results: list[dict[str, Any]] = []
    # XmlEngine addresses ``document.tables``.  python-docx exposes only
    # top-level tables there, so nested tables must not consume a table index.
    # Their text remains covered by the XML paragraph anchors above.
    for current, table in enumerate(document.tables):
        # Keep the elements themselves alive.  Storing ``id(cell._tc)`` alone
        # lets a short-lived wrapper be garbage-collected and its id reused by
        # a later row, which incorrectly hides real rows after a merged cell.
        seen_cells: set[Any] = set()
        for row_no, row in enumerate(table.rows):
            prior_label = ""
            logical_col = -1
            for col_no, cell in enumerate(row.cells):
                identity = cell._tc
                if identity in seen_cells:
                    continue
                seen_cells.add(identity)
                logical_col += 1
                value = cell.text.strip()
                if value:
                    prior_label = _short_table_label(value) or ""
                elif prior_label and col_no > 0:
                        # Never convert an arbitrary empty data-table cell into a
                        # field: labels must look like a short label or end in colon.
                    if prior_label:
                        # ``col`` is the python-docx / XmlEngine grid coordinate.
                        # Keep the compact merged-cell coordinate outside target:
                        # executors must receive only their supported schema.
                        target = {"kind": "cell", "table": current, "row": row_no, "col": col_no}
                        identity_text = json.dumps(target, sort_keys=True, separators=(",", ":"))
                        protected, reason = _protected("", prior_label)
                        results.append({"id": hashlib.sha256(identity_text.encode()).hexdigest()[:20],
                                        "label": prior_label, "context": {"previous": "", "paragraph": prior_label,
                                        "next": "相邻空表格单元格"}, "target": target, "protected": protected,
                                        "protected_reason": reason,
                                        "logical_location": f"table[{current}].r[{row_no}].c[{logical_col}]"})
                        # Only the immediately adjacent empty cell is a safe
                        # interpretation of a label.  Later cells could belong
                        # to a repeated table/data-row layout.
                    prior_label = ""
    return results


def _word_slots(path: Path) -> list[dict[str, Any]]:
    with zipfile.ZipFile(path) as package:
        parts = sorted(name for name in package.namelist() if _PART_RE.fullmatch(name))
        slots = []
        for part in parts:
            xml = package.read(part)
            slots.extend(_paragraph_slots(part, scan_paragraphs(xml), xml))
    slots.extend(_word_cell_slots(path))
    return slots


def _excel_label(value: str) -> str:
    return value.strip().lstrip("*＊").rstrip("：:").strip() or "待确认空位"


def _xlsx_slots(path: Path) -> list[dict[str, Any]]:
    import openpyxl

    book = openpyxl.load_workbook(path, read_only=False, data_only=False)
    results: list[dict[str, Any]] = []
    try:
        for sheet in book.worksheets:
            for cell in list(sheet._cells.values()):
                if not isinstance(cell.value, str):
                    continue
                raw = cell.value
                context = {"previous": "", "paragraph": _bounded(raw), "next": f"工作表 {sheet.title}"}
                spans = set(m.span() for m in _UNDERLINE_RE.finditer(raw)) | set(m.span() for m in _BRACKET_RE.finditer(raw))
                for match in _SUFFIX_RE.finditer(raw):
                    if _PLACEHOLDER_PREFIX_RE.search(raw[max(0, match.start() - 28):match.start()]): spans.add(match.span("slot"))
                for match in _COLON_END_RE.finditer(raw):
                    if not _SECTION_LABEL_RE.search(_label(raw, match.start())): spans.add((match.end(), match.end()))
                for start, end in _normalise_spans(spans):
                    label = _label(raw, start)
                    protected, reason = _protected(raw, label, start, end)
                    # XlsxEngine addresses a cell and a character span directly;
                    # it cannot execute Word's part/paragraph anchor protocol.
                    target = {"kind": "xlsx_cell", "sheet": sheet.title, "cell": cell.coordinate,
                              "expected_text": raw, "span_start": start, "span_end": end}
                    key = json.dumps(target, ensure_ascii=False, sort_keys=True, separators=(",", ":"))
                    results.append({"id": hashlib.sha256(key.encode()).hexdigest()[:20], "label": label, "context": context,
                                    "target": target, "protected": protected, "protected_reason": reason})
                # Adjacent blanks are considered only for an explicit label; this
                # deliberately does not invent rows for ordinary spreadsheet data.
                if raw.rstrip().endswith(("：", ":")):
                    merged = next((r for r in sheet.merged_cells.ranges if cell.coordinate in r), None)
                    next_col = merged.max_col + 1 if merged else cell.column + 1
                    right = sheet.cell(cell.row, next_col)
                    if right.value in (None, "") and right.data_type != "f" and right.has_style:
                        target = {"kind": "xlsx_cell", "sheet": sheet.title, "cell": right.coordinate,
                                  "label_cell": cell.coordinate}
                        key = json.dumps(target, ensure_ascii=False, sort_keys=True, separators=(",", ":"))
                        label = _excel_label(raw)
                        protected, reason = _protected(raw, label)
                        results.append({"id": hashlib.sha256(key.encode()).hexdigest()[:20], "label": label, "context": context,
                                        "target": target, "protected": protected, "protected_reason": reason})
    finally:
        book.close()
    return results


def discover_slots(path: Path | str) -> list[dict[str, Any]]:
    """Return unique, addressable template blanks without assigning any value.

    Excel data rows without an explicit nearby label are intentionally unsupported;
    creating new row/column structure is outside template slot discovery.
    """
    target = Path(path)
    if target.suffix.lower() == ".docx": slots = _word_slots(target)
    elif target.suffix.lower() == ".xlsx": slots = _xlsx_slots(target)
    else: raise ValueError(f"Only .docx and .xlsx are supported: {target}")
    unique: dict[str, dict[str, Any]] = {}
    occupied: set[tuple[str, int, int, int]] = set()
    for slot in slots:
        t = slot["target"]
        if t["kind"] == "anchor":
            # Excel anchors share a synthetic paragraph index, so their cell is
            # part of the overlap scope.  Word locations have no ``cell`` key.
            key = (str(t["part"]) + "!" + str(t.get("cell", "")), int(t["paragraph_index"]),
                   int(t["span_start"]), int(t["span_end"]))
            if any(key[0] == p and key[1] == i and not (key[3] <= s or key[2] >= e)
                   for p, i, s, e in occupied if not (s == e == key[2] == key[3])):
                continue
            occupied.add(key)
        elif t["kind"] == "xlsx_cell" and "expected_text" in t:
            key = (str(t["sheet"]) + "!" + str(t["cell"]), 0, int(t["span_start"]), int(t["span_end"]))
            if any(key[0] == p and not (key[3] <= s or key[2] >= e)
                   for p, _i, s, e in occupied if not (s == e == key[2] == key[3])):
                continue
            occupied.add(key)
        unique[slot["id"]] = slot
    return list(unique.values())
