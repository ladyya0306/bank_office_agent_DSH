"""Field-aware document filling for loan/credit dossiers.

The problem this solves: the same small set of facts (borrower, legal rep,
contract numbers, credit line) has to be typed into a dozen Word forms, and the
forms must come out **format-identical** - same fonts, underlines, table
borders, header/footer, pagination.

Strategy: never rebuild a document. Locate the *existing* text span that holds a
value (by label anchor, or by table label cell) and rewrite only its characters.
Word's character/paragraph properties live on the run and paragraph, so swapping
text inside an existing run leaves every formatting attribute untouched.

Two things this deliberately refuses to guess:
  * a field whose source value is unknown -> left blank, listed for review;
  * a field that matches more than one place and whose rule only permits one
    -> left untouched, listed for review.
Both show up in the review report instead of becoming a silent wrong document.
"""
from __future__ import annotations

import csv
import json
import re
import zipfile
from pathlib import Path
from typing import Any, Iterable, Iterator

from lxml import etree

from .common import (
    OfficeKitError,
    Result,
    ensure_parent,
    out_dir,
    resolve_inputs,
    safe_stem,
    unique_path,
    write_text,
)

# --------------------------------------------------------------------------
# profile handling
# --------------------------------------------------------------------------
MISSING_MARKERS = {"", "none", "null", "n/a", "待补充", "缺失", "todo", "tbd"}


def load_profile(path: str | Path) -> dict[str, dict[str, Any]]:
    """Load the field dictionary, keeping per-field provenance metadata.

    A field may declare ``template`` instead of ``value``, e.g.
    ``{"template": "{联系人}  {联系电话}"}``, so one form field that expects two
    facts (a name and a phone on the same line) can be composed without changing
    how the underlying facts are stored.
    """
    p = Path(path)
    if not p.exists():
        raise OfficeKitError(f"profile not found: {p}")
    raw = json.loads(p.read_text(encoding="utf-8"))
    fields: dict[str, dict[str, Any]] = {}
    for key, val in raw.items():
        if key.startswith("_"):
            continue  # documentation entries such as _说明
        if isinstance(val, dict):
            entry = dict(val)
        else:
            entry = {"value": val}
        # Profiles in the wild use either key for provenance; normalise so
        # downstream stages (role detection, reports) always see `origin`.
        if not entry.get("origin") and entry.get("source"):
            entry["origin"] = entry["source"]
        v = entry.get("value")
        entry["missing"] = v is None or (isinstance(v, str) and v.strip().lower() in MISSING_MARKERS)
        entry["key"] = key
        fields[key] = entry

    # Resolve composed fields, two passes so a template may reference another.
    import re as _re

    def render(tpl: str) -> tuple[str, list[str]]:
        missing: list[str] = []

        def sub(m):
            ref = m.group(1).strip()
            src = fields.get(ref)
            if src is None or src.get("missing"):
                missing.append(ref)
                return ""
            return str(src["value"])

        return _re.sub(r"\{([^}]+)\}", sub, tpl), missing

    for _ in range(2):
        for entry in fields.values():
            tpl = entry.get("template")
            if not tpl:
                continue
            text, missing = render(tpl)
            entry["value"] = text
            entry["missing"] = bool(missing) or not text.strip()
            if missing:
                entry["missing_refs"] = missing
    return fields


def profile_value(fields: dict[str, dict[str, Any]], key: str | None) -> tuple[str | None, dict[str, Any] | None]:
    if not key:
        return None, None
    f = fields.get(key)
    if f is None:
        return None, None
    if f.get("missing"):
        return None, f
    return str(f["value"]), f


# --------------------------------------------------------------------------
# run-level text surgery
# --------------------------------------------------------------------------
def paragraph_text(par) -> str:
    """Text of a paragraph, including any text inside hyperlink runs."""
    return "".join(r.text for r in par.runs)


def _first_format_run(par, start: int):
    """The run that owns character index ``start`` within the paragraph.

    The interesting case is a value that begins exactly where a run begins: the
    previous run *ends* at that offset but holds unrelated formatting (often the
    plain label), so returning it would copy the wrong character properties. The
    run that actually contains the offset wins.
    """
    pos = 0
    for r in par.runs:
        r_start, r_end = pos, pos + len(r.text)
        pos = r_end
        if r_start <= start < r_end:
            return r
    # start is at/after the end of the paragraph: fall back to the last run that
    # has text, which is what an append should inherit from.
    for r in reversed(par.runs):
        if r.text:
            return r
    return par.runs[0] if par.runs else None


def _copy_run_format(src, dst) -> None:
    """Copy character formatting (font, size, bold, underline, colour) run->run."""
    if src is None or dst is None:
        return
    try:
        dst.bold = src.bold
        dst.italic = src.italic
        dst.underline = src.underline
    except Exception:  # noqa: BLE001
        pass
    try:
        if src.font is not None:
            dst.font.name = src.font.name
            dst.font.size = src.font.size
            dst.font.color.rgb = src.font.color.rgb
    except Exception:  # noqa: BLE001
        pass
    try:
        from docx.oxml.ns import qn

        src_rpr = src._element.find(qn("w:rPr"))
        if src_rpr is not None:
            dst_rpr = dst._element.find(qn("w:rPr"))
            if dst_rpr is not None:
                dst._element.remove(dst_rpr)
            import copy as _copy

            dst._element.insert(0, _copy.deepcopy(src_rpr))
    except Exception:  # noqa: BLE001
        pass


def replace_in_paragraph(par, start: int, end: int, new_text: str) -> bool:
    """Replace characters [start, end) of a paragraph, preserving formatting.

    Only the *text* of runs is modified; run elements and their ``rPr`` are left
    completely untouched. That is what makes this format-neutral: an existing run
    already carries the font, size and underline the template author chose, so
    copying properties around would only risk overwriting them.

    (An earlier version copied the first overlapping run's ``rPr`` onto the
    receiving run. On a real form that silently rewrote ``hAnsi`` from 华文楷体 to
    方正宋三简体 - a genuine formatting change introduced by the tool.)
    """
    if start < 0 or end < start:
        return False

    if not par.runs:
        # A genuinely empty paragraph/cell has no run to write into: create one.
        if new_text:
            par.add_run(new_text)
            return True
        return False

    pos = 0
    done_insert = False
    for r in par.runs:
        r_start, r_end = pos, pos + len(r.text)
        pos = r_end
        if end > start:
            # Replacement: skip runs that do not overlap [start, end).
            if r_end <= start or r_start >= end:
                continue
        else:
            # Pure insertion at `start` belongs to the run ending at/after it.
            # `r_end <= start` would wrongly skip the run the text appends to when
            # start equals its length.
            if r_end < start or r_start > start:
                continue
        keep_head = r.text[: max(0, start - r_start)] if r_start < start else ""
        keep_tail = r.text[max(0, end - r_start):] if r_end > end else ""
        if not done_insert:
            r.text = keep_head + new_text + keep_tail
            done_insert = True
        else:
            r.text = keep_head + keep_tail
    return done_insert


def iter_paragraphs(document) -> Iterable[tuple[str, Any]]:
    """Yield (location_label, paragraph) across body, tables, headers, footers.

    Nested tables are walked too - loan forms embed tables inside table cells.
    """
    def walk_table(tbl, prefix: str):
        for ri, row in enumerate(tbl.rows):
            for ci, cell in enumerate(row.cells):
                loc = f"{prefix}/table[{ri},{ci}]"
                for par in cell.paragraphs:
                    yield loc, par
                for nested in cell.tables:
                    yield from walk_table(nested, loc)

    for i, par in enumerate(document.paragraphs):
        yield f"body[{i}]", par
    for ti, tbl in enumerate(document.tables):
        yield from walk_table(tbl, f"table{ti}")
    for si, section in enumerate(document.sections, start=1):
        for kind, part in (("header", section.header), ("footer", section.footer)):
            if part is None:
                continue
            for i, par in enumerate(part.paragraphs):
                yield f"{kind}{si}[{i}]", par
            for ti, tbl in enumerate(part.tables):
                yield from walk_table(tbl, f"{kind}{si}/table{ti}")


# --------------------------------------------------------------------------
# individual fill rules
# --------------------------------------------------------------------------
def _paragraph_offsets(root) -> Iterator[tuple[Any, int]]:
    """Yield (paragraph_element, byte offset) for every <w:p> under ``root``.

    The offset is the cumulative serialised length of every earlier sibling of the
    paragraph's parent, so it is the paragraph's true source position regardless of
    nesting. The walk is generic on purpose: these forms embed content in text
    boxes (``w:txbxContent``) *inside* table cells, and with 183 paragraphs of
    which 50 live in text boxes, any fixed cell/row traversal silently loses them
    and the python-docx/XML pairing collapses.
    """
    if root is None:
        return
    for el in root:
        if el.tag == qn_p():
            pos = 0
            parent = el.getparent()
            for sib in parent:
                if sib is el:
                    break
                pos += len(etree.tostring(sib))
            yield el, pos
        else:
            yield from _paragraph_offsets(el)


def _xml_index_by_offset(document) -> list[tuple[Any, Any]]:
    """Every body <w:p> as (element, paragraph) in document order.

    Pairing with the raw-XML scan is by sequence, using lxml's own paragraph
    iterator. Position alone is not reliable: these forms contain
    ``mc:AlternateContent`` blocks, and the parser's view of a Choice/Fallback pair
    differs from a byte-level scan, so the two sequences can drift by a few
    paragraphs. Sequence alignment absorbs that drift instead of mis-targeting a
    write (which is what a naive zip did).
    """
    from docx.text.paragraph import Paragraph

    body = document.element.body
    result: list[tuple[Any, Any]] = []

    # Cache a python-docx paragraph per element so identity comparisons work.
    for el in body.iter(qn_p()):
        parent = el.getparent()
        par = None
        try:
            if parent is not None and parent.tag != qn_body():
                from docx.table import _Cell

                par = Paragraph(el, _Cell(parent, document))
        except Exception:  # noqa: BLE001
            par = None
        if par is None:
            par = Paragraph(el, document)
        result.append((el, par))
    return result


def qn_body():
    from docx.oxml.ns import qn

    return qn("w:body")


def qn_p():
    from docx.oxml.ns import qn

    return qn("w:p")


def qn_tbl():
    from docx.oxml.ns import qn

    return qn("w:tbl")


W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _ordered_pydocx_paragraphs(document) -> list[Any]:
    """Superseded by offset-based pairing; kept for callers that only need order."""
    return [par for _, par in _xml_index_by_offset(document)]


class XmlEngine:
    """Splice existing body/header/footer text; never re-serialise the package."""

    def __init__(self, template: Path):
        import docx

        from .xml_fill import scan_paragraphs

        self.template = Path(template)
        with zipfile.ZipFile(self.template) as z:
            self.xml = z.read("word/document.xml")
        self.paragraphs = scan_paragraphs(self.xml)
        self.document = docx.Document(str(template))
        self.edits: list[tuple[int, int, str]] = []
        self._part_xml = {"word/document.xml": self.xml}
        self._part_edits = {"word/document.xml": self.edits}
        self._part_by_el = {}

        # Both views walk the same document.xml in document order. Pair by that
        # position only when the paragraph counts and visible text agree. A fuzzy
        # text alignment can silently pair the wrong one of two repeated labels.
        self._entries = _xml_index_by_offset(self.document)
        self._by_el: dict[Any, Any] = {}
        # Uneditable paragraphs (for example field-code text) still occupy a
        # position in scan_paragraphs. Never compress subsequent indexes.
        self._part_by_el.update({el: 'word/document.xml' for el, _ in self._entries})
        if len(self._entries) == len(self.paragraphs):
            for (el, par), raw in zip(self._entries, self.paragraphs):
                # w:tab and w:br are not w:t character data.  They do not
                # affect offsets inside the text records we can actually edit.
                parsed = paragraph_text(par).replace("\t", "").replace("\n", "")
                if parsed == raw.text:
                    self._by_el[el] = raw
                    self._part_by_el[el] = "word/document.xml"

        # Read only parts already present in the zip. Accessing section.header
        # via python-docx can create new parts, including for linked sections.
        from types import SimpleNamespace
        with zipfile.ZipFile(self.template) as z:
            for name in z.namelist():
                if not re.fullmatch(r"word/(?:header|footer)[0-9]+\.xml", name):
                    continue
                raw_xml = z.read(name)
                self._part_xml[name] = raw_xml
                self._part_edits[name] = []
                for raw in scan_paragraphs(raw_xml):
                    element = object()
                    paragraph = SimpleNamespace(_element=element)
                    self._entries.append((element, paragraph))
                    self._by_el[element] = raw
                    self._part_by_el[element] = name

    def xml_for(self, par):
        return self._by_el.get(par._element)

    def _insert_into_empty(self, par, x, new_text: str) -> None:
        """Insert a run into a paragraph that has no <w:t> (an empty table cell).

        This is the only case where new markup is added; it carries the run
        properties Word already stored on the paragraph, so the inserted value
        picks up the cell's intended font instead of a default.

        ⚠️ 空单元格在 Word 里常常写成**自闭合的 `<w:p/>`**，它连 `</w:p>` 都没有。
        这时不能"往 `</w:p>` 前面插"，必须把这个开标签补成成对的
        `<w:p …>…</w:p>`，否则算出来的插入位置是错的（甚至为负），
        整个单元格的值就丢了。
        """
        frag = self._run_fragment(par, new_text)
        if x.self_closing:
            # `<w:p …/>` → `<w:p …>`：去掉结尾的 `/`，**保留 `>`**。
            # （早先写成 `open_tag[:-2]`，把 `>` 一起吃掉了，产出 `<w:p<w:r>…` 这种坏 XML。）
            raw = x.open_tag
            open_tag = (raw[:-2] + b">") if raw.endswith(b"/>") else raw
            self._part_edits[self._part_by_el[par._element]].append((x.start, x.end,
                               open_tag.decode("utf-8", "replace") + frag + "</w:p>"))
            return
        at = x.end - len(b"</w:p>")
        self._part_edits[self._part_by_el[par._element]].append((at, at, frag))

    def fill_span(self, par, start: int, end: int, new_text: str) -> bool:
        from .xml_fill import splice
        from .value_fit import fit_value

        x = self.xml_for(par)
        if x is None:
            # No raw-XML counterpart (e.g. inside an mc:AlternateContent region
            # the parser and the byte scan disagree about). Refuse rather than
            # reporting a write that never happened.
            return False
        if not x.records:
            self._insert_into_empty(par, x, new_text)
            return True
        new_text = fit_value(x.text, start, end, new_text)
        edits = splice(x.records, start, end, new_text)
        if not edits:
            # Nothing overlapped the requested span: the value would be lost, so
            # this must surface as a failure instead of a silent success.
            return False
        for byte_start, byte_end, replacement in edits:
            self._part_edits[self._part_by_el[par._element]].append((byte_start, byte_end, replacement))
        return True

    def _run_fragment(self, par, text: str) -> str:
        """Insert text into an empty paragraph using its existing run defaults.

        The paragraph's pPr/rPr already supplies the font. Copying it into a
        new run makes lxml add namespace declarations and changes OOXML beyond
        the text slot, which failed the template-preservation check.
        """
        from .xml_fill import _escape
        return f'<w:r><w:t xml:space="preserve">{_escape(text)}</w:t></w:r>'

    def fill_cell(self, spec: dict[str, Any], new_value: str) -> bool:
        t_idx, r, c = int(spec["table"]), int(spec["row"]), int(spec["col"])
        try:
            cell = self.document.tables[t_idx].cell(r, c)
        except IndexError as exc:
            raise OfficeKitError(f"cell ({r},{c}) outside table {t_idx}") from exc
        target = cell.paragraphs[0]
        raw_target = self.xml_for(target)
        length = len(raw_target.text) if raw_target else len(paragraph_text(target))
        ok = self.fill_span(target, 0, length, new_value)
        # Clear any further paragraphs in the cell so stale text cannot linger.
        for extra in cell.paragraphs[1:]:
            raw_extra = self.xml_for(extra)
            length = len(raw_extra.text) if raw_extra else len(paragraph_text(extra))
            ok = self.fill_span(extra, 0, length, "") or ok
        return ok

    def find_anchor_hits(self, spec: dict[str, Any]) -> list[tuple[Any, int, int]]:
        from .xml_fill import find_targets

        required = spec.get("required")
        required_re = re.compile(required) if required else None
        part = spec.get("part", "word/document.xml")
        if part not in self._part_xml:
            raise OfficeKitError(f"模板不存在文字部件：{part}")

        allowed = None
        position = ("table", "row", "col")
        if any(k in spec for k in position):
            if part != "word/document.xml":
                raise OfficeKitError("页眉页脚暂只支持指定文字位置，不使用正文表格坐标")
            if not all(k in spec for k in position):
                raise OfficeKitError("Word 表格位置须同时给出 table、row、col")
            t_idx, row, col = (int(spec[k]) for k in position)
            try:
                table = self.document.tables[t_idx]
                cell = table.cell(row, col)
            except IndexError as exc:
                raise OfficeKitError(
                    f"Word 表格位置不存在：table={t_idx}, row={row}, col={col}"
                ) from exc
            allowed = set(cell._tc.iter(qn_p()))

            # Many bank forms put the printed label in one merged cell and the
            # blank value in the next cell. The rule's coordinates name the
            # value cell; its anchor is a check on the left-hand label.
            anchor = spec.get("anchor")
            if anchor and anchor not in cell.text and not cell.text.strip():
                label_cells = set()
                for left_col in range(col):
                    neighbour = table.cell(row, left_col)
                    if anchor in neighbour.text:
                        label_cells.add(neighbour._tc)
                if len(label_cells) == 1 and cell.paragraphs:
                    target = cell.paragraphs[0]
                    raw = self.xml_for(target)
                    if raw is not None:
                        return [(target, 0, len(raw.text))]
                # A missing or repeated label cannot justify writing into an
                # otherwise anonymous blank cell.
                return []

        hits: list[tuple[Any, int, int]] = []
        # paragraph_index is scoped to the selected XML part and follows the
        # same order as scan_paragraphs (including table paragraphs).  It is a
        # stable, reviewable locator when the complete paragraph text is also
        # supplied; it is intentionally not a document-global byte offset.
        paragraph_index = spec.get("paragraph_index")
        if paragraph_index is not None:
            try:
                paragraph_index = int(paragraph_index)
            except (TypeError, ValueError) as exc:
                raise OfficeKitError("paragraph_index 必须是整数") from exc
            if paragraph_index < 0:
                raise OfficeKitError("paragraph_index 不能为负数")
        part_index = 0
        for el, par in self._entries:
            if self._part_by_el.get(el) != part:
                continue
            current_index = part_index
            part_index += 1
            if paragraph_index is not None and current_index != paragraph_index:
                continue
            if allowed is not None and el not in allowed:
                continue
            entry = self._by_el.get(el)
            if entry is None:
                continue
            text = entry.text
            if "expected_text" in spec and str(spec["expected_text"]) != text:
                raise OfficeKitError(
                    "指定段落原文不匹配：paragraph_index=%s" % paragraph_index)
            if not text:
                continue
            if required_re is not None and not required_re.search(text):
                continue
            for s, e in find_targets(spec, text):
                hits.append((par, s, e))
        occurrence = spec.get("occurrence")
        if occurrence is not None:
            try:
                occurrence = int(occurrence)
            except (TypeError, ValueError) as exc:
                raise OfficeKitError("occurrence 必须是从 1 开始的整数") from exc
            if occurrence < 1:
                raise OfficeKitError("occurrence 必须从 1 开始")
            if len(hits) < occurrence:
                raise OfficeKitError(
                    "指定 occurrence 不存在：需要第 %d 个，实际 %d 个" %
                    (occurrence, len(hits)))
            hits = [hits[occurrence - 1]]
        if paragraph_index is not None and part_index <= paragraph_index:
            raise OfficeKitError(
                "指定段落不存在：part=%s, paragraph_index=%d" %
                (part, paragraph_index))
        return hits

    def save(self, dst: Path) -> Path:
        from .xml_fill import apply_edits, rewrite_zip

        changed = {name: apply_edits(self._part_xml[name], edits)
                   for name, edits in self._part_edits.items() if edits}
        rewrite_zip(self.template, Path(dst), changed.get("word/document.xml", self.xml),
                    extra_parts={name: data for name, data in changed.items()
                                 if name != "word/document.xml"})
        return Path(dst)


def save_pydocx_with_repair(document, template: Path, dst: Path, res: Result | None = None) -> None:
    """Save a python-docx document and undo the non-text damage it causes.

    python-docx re-serialises the package: it materialises empty header/footer
    parts, adds matching ``sectPr`` references, and drops ``xml:space``. Left
    alone, the first two produce a file with a dangling relationship that Word and
    python-docx both refuse to open.

    This lives in one place on purpose. It was previously duplicated in the
    fillmap and db-fill paths, and the second copy silently lost the xml:space
    step - so the same corrupt-package bug came back.
    """
    ensure_parent(dst)
    document.save(str(dst))
    from .xml_fill import preserve_space, repair_package

    repair = repair_package(template, dst)
    if res is not None and (repair.get("removed_parts")
                            or repair.get("removed_section_references")):
        res.data.setdefault("fidelity_repairs", {})[template.name] = repair

    with zipfile.ZipFile(dst) as z:
        names = z.namelist()
        payload = {n: z.read(n) for n in names}
    if "word/document.xml" not in payload:
        return
    fixed = preserve_space(payload["word/document.xml"])
    if fixed == payload["word/document.xml"]:
        return
    tmp = dst.with_suffix(dst.suffix + ".fix")
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for n in names:
            zout.writestr(n, fixed if n == "word/document.xml" else payload[n])
    tmp.replace(dst)


class XlsxEngine:
    """Fill a workbook by writing only the addressed cells.

    openpyxl round-trips the whole workbook on save, so this cannot claim byte
    fidelity the way the XML engine does. It does preserve cell styles, number
    formats, merged ranges and sheet-level settings. A non-empty target may be
    filled only when its rule supplies a unique label anchor and the text after
    that anchor is blank or an explicit-looking placeholder; existing values
    are never overwritten.
    """

    def __init__(self, template: Path):
        import openpyxl

        self.template = Path(template)
        self.wb = openpyxl.load_workbook(str(template))
        self.edits: list[tuple[str, str, str]] = []
        # Precise spans are expressed against the original cell text.  Keep
        # one immutable baseline and rebuild it after every edit so a longer
        # replacement cannot shift the next span's coordinates.
        self._precise: dict[tuple[str, str], dict[str, Any]] = {}

    @staticmethod
    def _adjacent_unit_value_cell(ws, cell):
        """Return the adjacent unit-wrapper value cell for a trailing-colon label."""
        current = cell.value
        if not isinstance(current, str) or not current.rstrip().endswith(("：", ":")):
            return None
        merged = next((r for r in ws.merged_cells.ranges if cell.coordinate in r), None)
        next_col = merged.max_col + 1 if merged else cell.column + 1
        candidate = ws.cell(cell.row, next_col)
        from .value_fit import unit_wrapper_span
        return candidate if unit_wrapper_span(candidate.value) else None

    def fill_xlsx_cell(self, spec: dict[str, Any], value: str) -> bool:
        coord = str(spec["cell"])
        sheet = spec.get("sheet")
        if sheet is None:
            ws = self.wb.worksheets[0]
        elif isinstance(sheet, int):
            try:
                ws = self.wb.worksheets[sheet]
            except IndexError as exc:
                raise OfficeKitError(f"工作表序号 {sheet} 不存在") from exc
        else:
            try:
                ws = self.wb[str(sheet)]
            except KeyError as exc:
                raise OfficeKitError(f"工作表 {sheet!r} 不存在") from exc

        # MergedCell objects cannot hold values. Always address the range's
        # top-left cell, while retaining the selected coordinate for errors.
        requested_coord = coord
        for merged in ws.merged_cells.ranges:
            if coord in merged:
                coord = ws.cell(merged.min_row, merged.min_col).coordinate
                break
        cell = ws[coord]
        current = cell.value
        if cell.data_type == "f" or (isinstance(current, str) and current.startswith("=")):
            raise OfficeKitError(f"{ws.title}!{requested_coord} 是公式单元格，禁止填充")
        adjacent_value = self._adjacent_unit_value_cell(ws, cell)
        if adjacent_value is not None:
            raise OfficeKitError(
                f"{ws.title}!{requested_coord} 是标签格，右侧 {adjacent_value.coordinate} 已有单位和空位；"
                "为避免重复写金额，请把字段映射到右侧值格"
            )
        if "span_start" in spec or "span_end" in spec:
            if not all(k in spec for k in ("expected_text", "span_start", "span_end")):
                raise OfficeKitError("xlsx 精确位置必须提供 expected_text、span_start、span_end")
            try:
                start, end = int(spec["span_start"]), int(spec["span_end"])
            except (TypeError, ValueError) as exc:
                raise OfficeKitError("xlsx span_start/span_end 必须是整数") from exc
            observed = "" if current is None else str(current)
            key = (ws.title, coord)
            state = self._precise.get(key)
            original = state["original"] if state is not None else observed
            if str(spec["expected_text"]) != original:
                raise OfficeKitError(f"{ws.title}!{requested_coord} 原文不匹配，拒绝修改")
            if start < 0 or end < start or end > len(original):
                raise OfficeKitError(f"{ws.title}!{requested_coord} 精确区间越界")
            blank_marks = "_＿-—.·…□☐"
            if any(ch not in blank_marks and not ch.isspace() for ch in original[start:end]):
                raise OfficeKitError(f"{ws.title}!{requested_coord} 精确区间含已有文字")
            state = self._precise.setdefault(key, {"original": original, "edits": []})
            if state["original"] != original:
                raise OfficeKitError(f"{ws.title}!{requested_coord} 原文已改变")
            if any(start < b and a < end for a, b, _ in state["edits"]):
                raise OfficeKitError(f"{ws.title}!{requested_coord} 精确区间重叠")
            from .value_fit import fit_value
            fitted = fit_value(original, start, end, value)
            state["edits"].append((start, end, fitted))
            rebuilt = original
            for a, b, replacement in sorted(state["edits"], reverse=True):
                rebuilt = rebuilt[:a] + replacement + rebuilt[b:]
            cell.value = rebuilt
            self.edits.append((ws.title, coord, str(cell.value)))
            return True
        if current in (None, "") or not str(current).strip():
            cell.value = value
        else:
            text = str(current)
            anchor = spec.get("anchor")
            if not anchor:
                raise OfficeKitError(
                    f"{ws.title}!{requested_coord} 已有内容，规则必须提供 anchor 才能安全追加"
                )
            anchor = str(anchor)
            if text.count(anchor) != 1:
                raise OfficeKitError(
                    f"{ws.title}!{requested_coord} 的 anchor 未唯一命中，拒绝修改"
                )
            start = text.index(anchor) + len(anchor)
            before = spec.get("before")
            if before:
                before = str(before)
                if text.count(before) != 1 or text.index(before) < start:
                    raise OfficeKitError(
                        f"{ws.title}!{requested_coord} 的 before 未在 anchor 后唯一命中，拒绝修改"
                    )
                end = text.index(before)
            else:
                end = len(text)
            gap = text[start:end]
            placeholder = spec.get("placeholder")
            if placeholder is not None:
                from .value_fit import fit_value
                placeholder = str(placeholder)
                if not placeholder or gap.count(placeholder) != 1:
                    raise OfficeKitError(
                        f"{ws.title}!{requested_coord} 的 placeholder 未唯一命中，拒绝修改"
                    )
                p = gap.index(placeholder)
                left, right = gap[:p], gap[p + len(placeholder):]
                if left.strip() or right.strip():
                    raise OfficeKitError(
                        f"{ws.title}!{requested_coord} 占位符附近含已有内容，拒绝覆盖"
                    )
                fitted = fit_value(text, start + p, start + p + len(placeholder), value)
                gap_replacement = left + fitted + right
            else:
                from .value_fit import fit_value
                fitted = fit_value(text, start, end, value)
                # Accept whitespace-only gaps and conventional blank marks.
                # Any other text is treated as a pre-existing business value.
                blank_marks = "_＿-—.·…□☐"
                mark_positions = [i for i, ch in enumerate(gap) if ch in blank_marks]
                residual = "".join(ch for ch in gap if ch not in blank_marks)
                if residual.strip():
                    raise OfficeKitError(
                        f"{ws.title}!{requested_coord} anchor 后已有内容，拒绝覆盖"
                    )
                if mark_positions:
                    left = gap[:mark_positions[0]]
                    right = gap[mark_positions[-1] + 1:]
                    if any(ch not in blank_marks and not ch.isspace() for ch in left + right):
                        raise OfficeKitError(
                            f"{ws.title}!{requested_coord} 占位符格式不明确，拒绝修改"
                        )
                    gap_replacement = left + fitted + right
                else:
                    # A label with no following text has an implicit blank slot.
                    gap_replacement = gap + fitted
            cell.value = text[:start] + gap_replacement + text[end:]
        self.edits.append((ws.title, coord, str(cell.value)))
        return True

    def save(self, dst: Path) -> Path:
        dst = Path(dst)
        dst.parent.mkdir(parents=True, exist_ok=True)
        self.wb.save(str(dst))
        return dst


def _apply_cell(document, spec: dict[str, Any], fields, new_value: str):
    """Fill a table cell identified by its label cell."""
    t_idx = int(spec["table"])
    r = int(spec["row"])
    c = int(spec["col"])
    try:
        tbl = document.tables[t_idx]
    except IndexError:
        raise OfficeKitError(f"table index {t_idx} does not exist")
    try:
        cell = tbl.cell(r, c)
    except IndexError:
        raise OfficeKitError(f"cell ({r},{c}) outside table {t_idx}")

    mode = spec.get("mode", "after_label")
    target_par = None
    start = end = 0

    if mode == "whole":
        # Replace the cell's entire text, formatted like its first run.
        target_par = cell.paragraphs[0]
        start, end = 0, len(paragraph_text(target_par))
    else:
        anchor = spec.get("anchor")
        pattern = spec.get("pattern")
        if not anchor and not pattern:
            raise OfficeKitError(f"cell rule for ({r},{c}) needs 'anchor' or 'pattern'")
        for par in cell.paragraphs:
            text = paragraph_text(par)
            if pattern:
                m = re.search(pattern, text)
                if m:
                    target_par, start, end = par, m.start(), m.end()
                    break
            elif anchor and anchor in text:
                i = text.index(anchor) + len(anchor)
                target_par = par
                start = i
                # Consume trailing blanks so we overwrite the placeholder gap.
                end = start
                while end < len(text) and text[end] in " 　\t":
                    end += 1
                if end == start:
                    end = start
                break
            else:
                continue
    if target_par is None:
        raise OfficeKitError(
            f"label not found in cell ({r},{c}) of table {t_idx}: "
            f"{spec.get('anchor') or spec.get('pattern')!r}"
        )
    if mode == "whole":
        if not replace_in_paragraph(target_par, start, end, new_value):
            raise OfficeKitError(f"could not write into cell ({r},{c}) of table {t_idx}")
        # Blank any further paragraphs in the cell so stale text can't linger.
        for extra in cell.paragraphs[1:]:
            replace_in_paragraph(extra, 0, len(paragraph_text(extra)), "")
    elif not replace_in_paragraph(target_par, start, end, new_value):
        raise OfficeKitError(f"could not write into cell ({r},{c}) of table {t_idx}")
    return target_par


def _blank_span(text: str, start: int) -> int:
    """End of the blank (space/ideographic-space/tab) run beginning at ``start``."""
    e = start
    while e < len(text) and text[e] in " 　\t":
        e += 1
    return e


def _label_spans(text: str, spec: dict[str, Any]) -> list[tuple[int, int]]:
    """Character spans to overwrite for one paragraph, per the rule spec.

    Forms, in increasing precision:
      pattern                       regex; the (?P<value>...) group (or whole match)
      anchor + prefix               the blank run immediately BEFORE the label
                                    (for "___ 同志（身份证号：" layouts)
      anchor + before               text between two labels - bounded and
                                    width-independent, the best fit for blank forms
      anchor (blank follows)        the blank run right after the label
      anchor (no blank)             from the label to end of paragraph (insert)
    """
    spans: list[tuple[int, int]] = []
    pattern = spec.get("pattern")
    anchor = spec.get("anchor")
    before = spec.get("before")
    prefix = bool(spec.get("prefix"))
    max_blank = int(spec.get("max_blank", 60))

    if pattern:
        for m in re.finditer(pattern, text):
            if "value" in (m.groupdict() or {}):
                spans.append((m.start("value"), m.end("value")))
            else:
                spans.append((m.start(), m.end()))
        return spans

    if not anchor:
        raise OfficeKitError("anchor rule needs 'anchor' or 'pattern'")

    idx = 0
    while True:
        i = text.find(anchor, idx)
        if i < 0:
            break
        if prefix:
            s = i
            while s > 0 and text[s - 1] in " 　\t":
                s -= 1
            if (i - s) > max_blank:
                idx = i + 1
                continue
            spans.append((s, i))
            idx = i + 1
            continue
        s = i + len(anchor)
        e = _blank_span(text, s)
        if (e - s) > max_blank:
            idx = i + 1
            continue
        if before:
            j = text.find(before, e)
            if j < 0:
                idx = i + 1
                continue
            e = j
        elif e == s:
            e = len(text)  # nothing blank: append right after the label
        spans.append((s, e))
        idx = i + 1
    return spans


def _apply_anchor(document, spec: dict[str, Any], fields, new_value: str):
    """Find the spans a rule targets and rewrite only their characters.

    ``spec`` keys:
      anchor / pattern   how to locate the label or the value
      before             closing label bounding the blank (see _label_spans)
      required           regex that must ALSO be present in the paragraph, used
                         to disambiguate a label that repeats (e.g. 企业名称
                         appears for both the borrower and a blank guarantor block)
      max_matches        how many matches are acceptable before deferring to a human
    """
    required = spec.get("required")
    max_matches = int(spec.get("max_matches", 1))
    required_re = re.compile(required) if required else None

    hits: list[tuple[str, Any, int, int]] = []
    for loc, par in iter_paragraphs(document):
        text = paragraph_text(par)
        if not text:
            continue
        if required_re is not None and not required_re.search(text):
            continue
        for s, e in _label_spans(text, spec):
            hits.append((loc, par, s, e))

    if not hits:
        return [], "label_not_found"
    if spec.get("only_first") and len(hits) > 1:
        # A form that repeats a label for extra signatories: the first block is
        # the one described by the source documents; later ones are left alone.
        hits = hits[:1]
    if len(hits) > max_matches:
        return [h[0] for h in hits], "ambiguous"
    for loc, par, s, e in hits:
        # Keep the fallback python-docx writer subject to the same unit/type
        # guard as XmlEngine.  Otherwise an XML-engine fallback could append a
        # date range directly before a printed "年".
        from .value_fit import fit_value
        new_value = fit_value(paragraph_text(par), s, e, new_value)
        if e <= s:
            # Empty span: insert at the anchor point.
            replace_in_paragraph(par, s, s, new_value)
        else:
            replace_in_paragraph(par, s, e, new_value)
    return [h[0] for h in hits], "filled"


# --------------------------------------------------------------------------
# command
# --------------------------------------------------------------------------
def cmd_fillmap(args) -> Result:
    res = Result("fillmap")
    fields = load_profile(args.profile)
    mapping_path = Path(args.mapping)
    if not mapping_path.exists():
        raise OfficeKitError(f"mapping not found: {mapping_path}")
    mapping = json.loads(mapping_path.read_text(encoding="utf-8"))

    templates = resolve_inputs(args.input)
    out = out_dir(getattr(args, "out", None), "fillmap")
    strict = getattr(args, "strict", True)

    import docx

    engine_name = (getattr(args, "engine", None) or "xml").lower()
    report: list[dict[str, Any]] = []
    review: list[dict[str, Any]] = []
    produced: list[str] = []

    for tpl in templates:
        doc_rules = mapping.get(tpl.name) or mapping.get(tpl.stem)
        if doc_rules is None:
            res.warn(f"no mapping entry for {tpl.name}; skipped")
            continue

        engine = None
        book = None
        if tpl.suffix.lower() in (".xlsx", ".xlsm"):
            # A workbook target: fill addressed cells rather than Word paragraphs.
            try:
                book = XlsxEngine(tpl)
            except Exception as exc:  # noqa: BLE001
                res.warn(f"{tpl.name}: 无法打开工作簿（{exc}）")
                continue
            document = None
        else:
            if engine_name == "xml":
                try:
                    engine = XmlEngine(tpl)
                    document = engine.document
                except OfficeKitError as exc:
                    res.warn(f"{tpl.name}: XML 引擎不可用（{exc}），改用 python-docx 引擎")
                    engine = None
            if engine is None:
                document = docx.Document(str(tpl))
        filled_fields: set[str] = set()

        for rule in doc_rules.get("fields", []):
            key = rule.get("field")
            label = rule.get("label") or rule.get("anchor") or rule.get("pattern")
            value, meta = profile_value(fields, key)

            if value is None:
                reason = "profile_key_missing" if meta is None else "value_missing_in_source"
                review.append(
                    {
                        "file": tpl.name, "field": key, "label": label,
                        "location": _describe_target(rule.get("target", rule)),
                        "status": reason,
                        "note": (meta or {}).get("note", "") if meta else "该字段未在字典中定义",
                    }
                )
                continue

            if not str(value).strip():
                # An empty string is a missing value, not a value. Writing it would
                # blank the template's own default and then report success.
                review.append(
                    {
                        "file": tpl.name, "field": key, "label": label,
                        "location": _describe_target(rule.get("target", rule)),
                        "status": "value_missing_in_source",
                        "note": "字段字典中的值为空，未改动原文，需人工确认",
                    }
                )
                continue

            try:
                target = rule.get("target") or {}
                if target.get("kind") == "xlsx_cell":
                    if book is None:
                        raise OfficeKitError("目标不是工作簿")
                    book.fill_xlsx_cell(target, value)
                    report.append({"file": tpl.name, "field": key, "label": label,
                                   "status": "filled",
                                   "locations": f"单元格 {target['cell']}", "value": value})
                elif target.get("kind") == "cell":
                    if engine is not None:
                        if not engine.fill_cell(target, value):
                            raise OfficeKitError(f"could not write cell {target}")
                    else:
                        _apply_cell(document, target, fields, value)
                    report.append({"file": tpl.name, "field": key, "label": label,
                                   "status": "filled", "locations": _describe_target(target),
                                   "value": value})
                elif engine is not None:
                    hits = engine.find_anchor_hits(target)
                    if not hits:
                        review.append({"file": tpl.name, "field": key, "label": label,
                                       "location": "-", "status": "label_not_found",
                                       "note": "模板中未找到该标签，可能表格结构已变"})
                    else:
                        if target.get("only_first") and len(hits) > 1:
                            hits = hits[:1]
                        if len(hits) > int(target.get("max_matches", 1)):
                            review.append({
                                "file": tpl.name, "field": key, "label": label,
                                "location": f"{len(hits)} matches",
                                "status": "ambiguous_multiple_matches",
                                "note": f"出现 {len(hits)} 处，已跳过以免填错；请人工确认"})
                        else:
                            for par, s, e in hits:
                                engine.fill_span(par, s, e, value)
                            report.append({"file": tpl.name, "field": key, "label": label,
                                           "status": "filled",
                                           "locations": f"{len(hits)} spot(s)",
                                           "value": value})
                else:
                    locs, status = _apply_anchor(document, target, fields, value)
                    if status == "filled":
                        report.append({"file": tpl.name, "field": key, "label": label,
                                       "status": "filled", "locations": ";".join(locs), "value": value})
                    elif status == "ambiguous":
                        review.append({"file": tpl.name, "field": key, "label": label,
                                       "location": f"{len(locs)} matches",
                                       "status": "ambiguous_multiple_matches",
                                       "note": f"出现 {len(locs)} 处，已跳过以免填错；请人工确认"})
                    else:
                        review.append({"file": tpl.name, "field": key, "label": label,
                                       "location": "-", "status": "label_not_found",
                                       "note": "模板中未找到该标签，可能表格结构已变"})
                filled_fields.add(key)
            except OfficeKitError as exc:
                review.append({"file": tpl.name, "field": key, "label": label, "location": "-",
                               "status": "rule_failed", "note": str(exc)})

        stem = safe_stem(tpl.stem)
        suffix = ".xlsx" if book is not None else ".docx"
        dst = unique_path(out / f"{stem}_已填写{suffix}")
        ensure_parent(dst)
        if book is not None:
            book.save(dst)
        elif engine is not None:
            engine.save(dst)
        else:
            save_pydocx_with_repair(document, tpl, dst, res)
        produced.append(str(dst))
        res.add_artifact(dst, f"{tpl.name} 填充结果")
        report.append({"file": tpl.name, "field": "-", "label": "-", "status": "output",
                       "locations": str(dst), "value": f"{len(filled_fields)} fields"})

    # ---- reports
    # Resolve each name against the filesystem ONCE, before writing anything.
    # unique_path() only inspects existing files, so calling it repeatedly before
    # any write returns the same path every time - which silently made four report
    # writes collide on one file, leaving only the last one on disk.
    p_csv = unique_path(out / "填充记录.csv")
    review_csv = unique_path(out / "待人工确认.csv") if review else None
    review_md = unique_path(out / "待人工确认.md") if review else None

    ensure_parent(p_csv)
    with open(p_csv, "w", encoding="utf-8-sig", newline="") as fh:
        w = csv.DictWriter(fh, fieldnames=["file", "field", "label", "status", "locations", "value"])
        w.writeheader()
        w.writerows(report)
    res.add_artifact(p_csv, "每个字段填了什么、填到哪")

    if review_csv is not None:
        ensure_parent(review_csv)
        with open(review_csv, "w", encoding="utf-8-sig", newline="") as fh:
            w = csv.DictWriter(fh, fieldnames=["file", "field", "label", "location", "status", "note"])
            w.writeheader()
            w.writerows(review)
        res.add_artifact(review_csv, "需要人工确认的字段")

        write_text(review_md, _review_markdown(review, report))
        res.add_artifact(review_md, "待确认清单（可读版）")
        res.warn(f"{len(review)} 个字段未填充，已列入待人工确认清单")

    n_filled = sum(1 for r in report if r["status"] == "filled")
    res.data.update(
        {
            "profile_fields": len(fields),
            "templates": len(produced),
            "filled_assignments": n_filled,
            "review_items": len(review),
            "outputs": produced,
            "review_breakdown": _breakdown(review),
        }
    )
    if strict and not produced:
        raise OfficeKitError("nothing was produced; check the mapping keys against the template names")
    return res


def _describe_target(target: dict[str, Any]) -> str:
    """Render a rule target as a short, human-readable location."""
    if not isinstance(target, dict):
        return str(target)
    if target.get("kind") == "cell":
        return f"表格{target.get('table')} 第{target.get('row')}行第{target.get('col')}列"
    pat = target.get("pattern")
    if pat:
        return f"按标签匹配: {pat}"
    anc = target.get("anchor")
    return f"标签: {anc}" if anc else json.dumps(target, ensure_ascii=False)


def _breakdown(review: list[dict[str, Any]]) -> dict[str, int]:
    out: dict[str, int] = {}
    for r in review:
        out[r["status"]] = out.get(r["status"], 0) + 1
    return out


def _review_markdown(review: list[dict[str, Any]], report: list[dict[str, Any]]) -> str:
    L = ["# 待人工确认清单", ""]
    L.append("以下字段**没有被自动填写**，请人工处理。已填写的字段不受影响。")
    L.append("")
    by_status = _breakdown(review)
    L.append("## 汇总")
    L.append("")
    L.append("| 原因 | 数量 | 含义 |")
    L.append("| --- | --- | --- |")
    meaning = {
        "value_missing_in_source": "源文件里就没有这个值（如利率、用途）——需你填写",
        "profile_key_missing": "字段字典里没定义这个键——需补充字典",
        "ambiguous_multiple_matches": "同一标签在文档里出现多处，程序拒绝猜——需人工确认填哪处",
        "label_not_found": "模板里找不到该标签，可能版式已改",
        "rule_failed": "填充规则执行失败（表格坐标越界等）",
    }
    for k, v in by_status.items():
        L.append(f"| {k} | {v} | {meaning.get(k, '')} |")
    L.append("")
    L.append("## 明细")
    L.append("")
    L.append("| 文件 | 字段 | 标签 | 位置 | 原因 | 说明 |")
    L.append("| --- | --- | --- | --- | --- | --- |")
    for r in review:
        L.append(
            f"| {r['file']} | `{r['field']}` | {r['label']} | {r['location']} | "
            f"{r['status']} | {r.get('note', '')} |"
        )
    L.append("")
    L.append("## 已自动填写")
    L.append("")
    L.append("| 文件 | 字段 | 值 | 位置 |")
    L.append("| --- | --- | --- | --- |")
    for r in report:
        if r["status"] == "filled":
            L.append(f"| {r['file']} | `{r['field']}` | {r['value']} | {r['locations']} |")
    return "\n".join(L)
