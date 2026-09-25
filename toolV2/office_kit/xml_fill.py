"""Byte-faithful form filling: patch word/document.xml in place, touch nothing else.

Why this exists
---------------
Writing the document back through python-docx is *not* format-neutral: it
re-serialises every part, refreshes docProps timestamps, drops
``xml:space="preserve"``, and materialises empty header/footer parts. Verified on
a real template, that meant new ``<w:headerReference>``/``<w:footerReference>``
elements appeared in the output - a genuine structural change nobody asked for.

This module instead edits the raw XML text:
  * every ``<w:t>`` character datum is located by scanning the original bytes;
  * only the records overlapping the intended field span are rewritten;
  * the rebuilt zip re-uses every other part's bytes verbatim, and reuses the
    original document.xml bytes too, minus the spliced character ranges.

Result: the only part whose bytes differ is word/document.xml, and within it the
only difference is the characters that make up the filled values. There is no
serialisation step that could normalise, reorder, or drop anything.
"""
from __future__ import annotations

import re
import shutil
import zipfile
from dataclasses import dataclass
from html import unescape as html_unescape
from pathlib import Path
from typing import Any, Iterator, Sequence

from .common import OfficeKitError

# <w:t> may carry attributes; require an explicit close so the match can never
# run past the end of the element and swallow following markup.
TEXT_RE = re.compile(rb"<w:t(?:\s[^>]*)?>(.*?)</w:t>", re.S)
P_OPEN_RE = re.compile(rb"<w:p(?=[\s>/])")
P_CLOSE_RE = re.compile(rb"</w:p>")


@dataclass
class TextRecord:
    """One <w:t> node: where its body bytes live and what text it holds."""

    start: int          # byte offset of the body inside document.xml
    end: int
    text: str
    # Byte boundaries for every logical character in ``text``.  XML entities
    # are one displayed character but occupy several source bytes, so using
    # ``len(text[:n].encode('utf-8'))`` would drift after ``&amp;`` or ``&#...;``.
    char_byte_offsets: tuple[int, ...]


@dataclass
class XmlParagraph:
    """A <w:p> element seen as a sequence of text records."""

    start: int
    end: int
    records: list[TextRecord]
    #: 原样的开标签字节（含结尾的 `>` 或 `/>`），自闭合时要靠它还原成成对标签
    open_tag: bytes = b"<w:p>"
    #: `<w:p/>` 这种自闭合写法。**空表格单元格就是这种**，非常常见。
    self_closing: bool = False

    @property
    def text(self) -> str:
        return "".join(r.text for r in self.records)


def _decode(raw: bytes) -> str:
    try:
        return html_unescape(raw.decode("utf-8"))
    except UnicodeDecodeError:
        return html_unescape(raw.decode("utf-8", "replace"))


def _char_byte_offsets(raw: bytes, text: str) -> tuple[int, ...]:
    """Map logical text positions to byte offsets in one escaped ``w:t`` body.

    ``word/document.xml`` stores display text as UTF-8 plus XML entities.  A
    source range such as ``&amp;`` must therefore map to one character boundary,
    not five.  Word emits the five predefined XML entities and numeric character
    references; accepting other named HTML entities here would make an invalid
    Word XML document appear addressable, so they intentionally remain literal.
    """
    offsets = [0]
    i = 0
    logical: list[str] = []
    while i < len(raw):
        if raw[i:i + 1] == b"&":
            semi = raw.find(b";", i + 1)
            token = raw[i:semi + 1] if semi >= 0 else b""
            decoded = _decode(token) if token else ""
            # Valid XML entity references decode to exactly one character here.
            if token and decoded != token.decode("utf-8", "replace") and len(decoded) == 1:
                logical.append(decoded)
                i = semi + 1
                offsets.append(i)
                continue
        lead = raw[i]
        width = 1 if lead < 0x80 else 2 if lead < 0xE0 else 3 if lead < 0xF0 else 4
        chunk = raw[i:i + width]
        decoded = _decode(chunk)
        # Invalid XML should not make offsets unsafe.  The decoded scanner text
        # still has one replacement character for this byte sequence.
        if not decoded:
            decoded = "\ufffd"
        logical.extend(decoded)
        i += max(1, len(chunk))
        offsets.extend([i] * len(decoded))
    if "".join(logical) != text:
        # This is a programmer/data-integrity error, not a span that can be
        # guessed safely.  Refusing is preferable to cutting XML at an unknown
        # boundary.
        raise OfficeKitError("w:t XML entity decoding does not match logical text")
    return tuple(offsets)


def scan_paragraphs(xml: bytes) -> list[XmlParagraph]:
    """Find every <w:p>...</w:p> span and the <w:t> records inside it.

    Paragraphs never nest, so pairing <w:p ...> with the next </w:p> is exact and
    avoids a real XML parser - which matters, because parsing and re-serialising
    is precisely the normalisation this module exists to avoid.

    ⚠️ **自闭合的 `<w:p/>` 必须单独认出来**：空表格单元格在 Word 里就长这样，
    而它没有 `</w:p>`。早先的版本对它也去找"下一个 `</w:p>`"，
    于是这个空段落**吞掉了后面所有段落**，段落配对整体错位——
    表现是"往空单元格里填值"永远报 `write_failed`，而填别处看起来正常。
    """
    paragraphs: list[XmlParagraph] = []
    for m in P_OPEN_RE.finditer(xml):
        start = m.start()
        gt = xml.find(b">", m.start())
        if gt < 0:
            continue
        open_tag = xml[start:gt + 1]
        if open_tag.endswith(b"/>"):
            paragraphs.append(XmlParagraph(start=start, end=gt + 1, records=[],
                                           open_tag=open_tag, self_closing=True))
            continue
        close = P_CLOSE_RE.search(xml, m.end())
        end = close.end() if close else len(xml)
        body = xml[m.end():end]
        records: list[TextRecord] = []
        for tm in TEXT_RE.finditer(body):
            body_span = tm.span(1)
            raw_text = tm.group(1)
            decoded = _decode(raw_text)
            records.append(
                TextRecord(
                    start=m.end() + body_span[0],
                    end=m.end() + body_span[1],
                    text=decoded,
                    char_byte_offsets=_char_byte_offsets(raw_text, decoded),
                )
            )
        paragraphs.append(XmlParagraph(start=start, end=end, records=records,
                                       open_tag=open_tag, self_closing=False))
    return paragraphs


def _escape(text: str) -> str:
    return (
        text.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
    )


def splice(records: list[TextRecord], start: int, end: int, new_text: str) -> list[tuple[int, int, str]]:
    """Translate a logical [start,end) span into per-record byte replacements.

    Returns document-relative byte ranges.  Each range covers only the selected
    characters, rather than the entire ``w:t`` body.  That makes two independent
    fields in one run composable in a single save.
    """
    edits: list[tuple[int, int, str]] = []
    pos = 0
    inserted = False
    for rec in records:
        r_start, r_end = pos, pos + len(rec.text)
        pos = r_end
        if end > start:
            if r_end <= start or r_start >= end:
                continue
        else:
            if r_end < start or r_start > start:
                continue
        local_start = max(0, start - r_start)
        local_end = min(len(rec.text), end - r_start) if end > start else local_start
        byte_start = rec.start + rec.char_byte_offsets[local_start]
        byte_end = rec.start + rec.char_byte_offsets[local_end]
        if not inserted:
            edits.append((byte_start, byte_end, _escape(new_text)))
            inserted = True
        else:
            # A cross-run replacement removes the covered suffix/prefix in
            # subsequent records after inserting the value in the first one.
            edits.append((byte_start, byte_end, ""))
    return edits


def find_targets(spec: dict[str, Any], text: str) -> list[tuple[int, int]]:
    """Locate the spans a rule targets inside one paragraph's text."""
    spans: list[tuple[int, int]] = []
    pattern = spec.get("pattern")
    anchor = spec.get("anchor")
    before = spec.get("before")
    prefix = bool(spec.get("prefix"))
    max_blank = int(spec.get("max_blank", 60))

    # A precise span is still handled by the same XML splice path.  The
    # caller must provide the complete paragraph text at the higher layer;
    # here we only enforce a bounded, blank-only range before any edit is
    # scheduled.  This makes a model-supplied range unable to overwrite prose.
    if "span_start" in spec or "span_end" in spec:
        if "span_start" not in spec or "span_end" not in spec:
            raise OfficeKitError("精确位置必须同时提供 span_start 和 span_end")
        try:
            start, end = int(spec["span_start"]), int(spec["span_end"])
        except (TypeError, ValueError) as exc:
            raise OfficeKitError("span_start/span_end 必须是整数") from exc
        if start < 0 or end < start or end > len(text):
            raise OfficeKitError(
                "精确位置超出段落范围：span=[%d,%d), 段落长度=%d" %
                (start, end, len(text)))
        blank_marks = "_＿-—.·…□☐"
        if any(ch not in blank_marks and not ch.isspace() for ch in text[start:end]):
            raise OfficeKitError("精确位置包含已有文字，拒绝覆盖")
        if anchor and str(anchor) not in text:
            raise OfficeKitError("精确位置的 anchor 不在指定段落中")
        return [(start, end)]

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
            if (i - s) <= max_blank:
                spans.append((s, i))
            idx = i + 1
            continue
        s = i + len(anchor)
        e = s
        while e < len(text) and text[e] in " 　\t":
            e += 1
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
            e = len(text)
        spans.append((s, e))
        idx = i + 1
    return spans


def apply_edits(xml: bytes, edits: Sequence[tuple[int, int, str]]) -> bytes:
    """Apply byte-range replacements, right to left so offsets stay valid."""
    ordered = sorted(edits, key=lambda e: (e[0], e[1]))
    previous: tuple[int, int, str] | None = None
    for edit in ordered:
        start, end, _replacement = edit
        if start < 0 or end < start or end > len(xml):
            raise OfficeKitError("XML edit range outside document.xml")
        if previous is not None:
            p_start, p_end, _ = previous
            # Two insertions at exactly the same location are also ambiguous:
            # their result otherwise depends on ordering in the caller.
            same_point = p_start == p_end == start == end
            if p_end > start or same_point:
                raise OfficeKitError("overlapping XML text edits; duplicate or intersecting fill targets")
        previous = edit
    out = xml
    for start, end, replacement in sorted(ordered, key=lambda e: -e[0]):
        out = out[:start] + replacement.encode("utf-8") + out[end:]
    return out


def rewrite_zip(src: Path, dst: Path, new_document: bytes,
                extra_parts: dict[str, bytes] | None = None) -> None:
    """Copy package bytes, replacing body and explicitly edited existing headers."""
    extra_parts = extra_parts or {}
    if any(not re.fullmatch(r"word/(?:header|footer)[0-9]+\.xml", name)
           for name in extra_parts):
        raise OfficeKitError("Only existing header/footer text parts can be replaced")
    with zipfile.ZipFile(src) as zin:
        if not set(extra_parts).issubset(zin.namelist()):
            raise OfficeKitError("Cannot create new header/footer parts while filling")
    dst.parent.mkdir(parents=True, exist_ok=True)
    tmp = dst.with_suffix(dst.suffix + ".tmp")
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "word/document.xml":
                data = new_document
            elif item.filename in extra_parts:
                data = extra_parts[item.filename]
            # Preserve the original entry name and compression choice.
            new_info = zipfile.ZipInfo(item.filename, date_time=item.date_time)
            new_info.compress_type = item.compress_type
            new_info.external_attr = item.external_attr
            zout.writestr(new_info, data)
    shutil.move(str(tmp), str(dst))


# --------------------------------------------------------------------------
# fidelity repair for documents written by python-docx
# --------------------------------------------------------------------------
HDR_REF_RE = re.compile(rb"<w:(?:header|footer)Reference\b[^>]*/>")
HDR_REF_FULL_RE = re.compile(rb"<w:(?:header|footer)Reference\b[^>]*>.*?</w:(?:header|footer)Reference>", re.S)


def _has_references(xml: bytes) -> bool:
    return bool(HDR_REF_RE.search(xml) or HDR_REF_FULL_RE.search(xml))


def _strip_rels_pointing_to(xml: bytes, prefixes: tuple[str, ...]) -> tuple[bytes, list[str]]:
    """Drop <Relationship> entries whose Target starts with one of ``prefixes``.

    Returns (new_xml, removed_rel_ids). The ids are needed to also remove the
    sectPr references that named them - a relationship removed without its
    reference, or a reference left pointing at a deleted part, both produce a
    package Word refuses to open.
    """
    try:
        from lxml import etree

        parser = etree.XMLParser(remove_blank_text=False)
        root = etree.fromstring(xml, parser=parser)
    except Exception:  # noqa: BLE001
        return xml, []
    removed: list[str] = []
    ns = root.nsmap.get(None, "")
    tag = f"{{{ns}}}Relationship" if ns else "Relationship"
    for rel in list(root):
        if rel.tag != tag:
            continue
        target = rel.get("Target") or ""
        if any(target.startswith(p) or target.startswith("/" + p) for p in prefixes):
            if rel.get("Id"):
                removed.append(rel.get("Id"))
            root.remove(rel)
    if not removed:
        return xml, []
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True), removed


def _strip_rels_for_parts(items: dict[str, bytes],
                          removed_parts: list[str]) -> tuple[dict[str, bytes], list[str]]:
    """Remove relationships pointing at deleted parts; return their ids."""
    stems = [name.rsplit("/", 1)[-1] for name in removed_parts]
    out = dict(items)
    if not stems:
        return out, []
    rel_ids: list[str] = []
    for name in list(out):
        if not name.endswith(".rels"):
            continue
        new_xml, removed = _strip_rels_pointing_to(out[name], tuple(stems))
        if removed:
            out[name] = new_xml
            rel_ids.extend(removed)
    return out, rel_ids


def _strip_references_by_id(doc_xml: bytes, rel_ids: list[str]) -> tuple[bytes, int]:
    """Remove <w:headerReference>/<w:footerReference> naming the given r:ids."""
    if not doc_xml or not rel_ids:
        return doc_xml, 0
    removed = 0
    for rid in rel_ids:
        pattern = re.compile(
            rb"<w:(?:header|footer)Reference\b[^>]*r:id=\""
            + re.escape(rid.encode()) + rb"\"[^>]*/>")
        doc_xml, n = pattern.subn(b"", doc_xml)
        removed += n
    return doc_xml, removed


def repair_package(original: Path, written: Path) -> dict[str, Any]:
    """Undo what python-docx changes about the package beyond the text.

    python-docx materialises empty ``header1.xml``/``footer1.xml`` parts and adds
    matching ``sectPr`` references even when the original had none. Empty parts
    render nothing, but the extra relationships make the package differ, and a
    removal done carelessly leaves dangling references that make the file
    unopenable. This function removes exactly the added-empty parts, exactly the
    relationships that pointed at them, and exactly the references naming those
    relationships - nothing else.
    """
    original_parts: dict[str, bytes]
    with zipfile.ZipFile(original) as z:
        original_parts = {i.filename: z.read(i.filename) for i in z.infolist()}

    with zipfile.ZipFile(written) as z:
        written_items = [(i.filename, z.read(i.filename)) for i in z.infolist()]

    added_empty = []
    for name, data in written_items:
        if name in original_parts:
            continue
        if name.startswith(("word/header", "word/footer")) and name.endswith(".xml"):
            # Only drop it when it carries no content of its own.
            if len(data) < 2000 and b"<w:t" not in data:
                added_empty.append(name)

    if not added_empty:
        return {"removed_parts": [], "removed_section_references": False, "removed_rels": []}

    payload = {n: d for n, d in written_items}
    for name in added_empty:
        payload.pop(name, None)
    # Delete only the parts identified as *added and empty*. An earlier version
    # dropped anything named header1.xml/footer1.xml unconditionally, destroying a
    # footer the template legitimately owned and leaving a dangling reference.
    payload, rel_ids = _strip_rels_for_parts(payload, list(added_empty))

    # Remove exactly the sectPr references that pointed at those relationships.
    # Blanking *all* references (the previous approach) also removed a legitimate
    # header/footer the original claimed.
    doc_xml, n_refs = _strip_references_by_id(payload.get("word/document.xml", b""), rel_ids)
    if doc_xml:
        payload["word/document.xml"] = doc_xml
    changed_sectpr = n_refs > 0

    # Preserve the original part order where possible so the package stays tidy.
    order = [n for n, _ in written_items if n in payload]
    order += [n for n in payload if n not in order]

    tmp = written.with_suffix(written.suffix + ".repair")
    with zipfile.ZipFile(written) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        info_by_name = {i.filename: i for i in zin.infolist()}
        for name in order:
            info = info_by_name.get(name)
            data = payload[name]
            if info is not None:
                zi = zipfile.ZipInfo(name, date_time=info.date_time)
                zi.compress_type = info.compress_type
                zi.external_attr = info.external_attr
            else:
                zi = zipfile.ZipInfo(name)
                zi.compress_type = zipfile.ZIP_DEFLATED
            zout.writestr(zi, data)
    shutil.move(str(tmp), str(written))

    removed_rels = []
    for name in payload:
        if name.endswith(".rels"):
            _, removed = _strip_rels_pointing_to(payload[name], tuple(added_empty))
            removed_rels.extend(removed)

    return {"removed_parts": added_empty, "removed_section_references": changed_sectpr,
            "removed_rels": removed_rels}


def preserve_space(xml: bytes) -> bytes:
    """Ensure xml:space='preserve' survives on <w:t> nodes with edge whitespace.

    python-docx drops the attribute when it re-serialises a run, which silently
    changes how leading/trailing spaces in a value are rendered.
    """
    def fix(m: re.Match) -> bytes:
        if b"xml:space" in (m.group(1) or b""):
            return m.group(0)
        body = m.group(2)
        if not re.search(rb"(^ )|( $)", body):
            return m.group(0)
        return b"<w:t" + (m.group(1) or b"") + b' xml:space="preserve">' + body + b"</w:t>"

    return re.sub(rb"<w:t(\s[^>]*)?>(.*?)</w:t>", fix, xml, flags=re.S)


def part_has_meaningful_difference(a: bytes, b: bytes) -> bool:
    """True when two XML parts differ structurally, not just in serialisation."""
    try:
        from lxml import etree

        parser = etree.XMLParser(remove_blank_text=False, huge_tree=True)
        ra, rb = etree.fromstring(a, parser), etree.fromstring(b, parser)
        na = etree.tostring(ra, method="c14n2", with_comments=False)
        nb = etree.tostring(rb, method="c14n2", with_comments=False)
        return na != nb
    except Exception:  # noqa: BLE001
        return a != b


def iter_document_paragraphs(xml: bytes) -> Iterator[tuple[int, XmlParagraph]]:
    for i, p in enumerate(scan_paragraphs(xml)):
        yield i, p


# --------------------------------------------------------------------------
# 保真证明（红线三：格式未变 / 打不开）
# --------------------------------------------------------------------------
#: 单元格值 / 文本节点：比对前把**文字**挖空，只留结构（属性保留）
_W_T_RE = re.compile(rb"(<w:t(?:\s[^>]*)?>)(.*?)(</w:t>)", re.S)
_V_RE = re.compile(rb"(<v(?:\s[^>]*)?>)(.*?)(</v>)", re.S)
_INLINE_T_RE = re.compile(rb"(<t(?:\s[^>]*)?>)(.*?)(</t>)", re.S)
#: "纯文字 run"：`<w:r>` 里除了可选的 `<w:rPr>` 与一个 `<w:t>` 什么都没有。
#: 它是**文字的容器**，不是排版。填空单元格时新增的就是这种 run。
_PURE_TEXT_RUN_RE = re.compile(
    rb"<w:r>(?:<w:rPr>.*?</w:rPr>)?<w:t(?:\s[^>]*)?>.*?</w:t></w:r>", re.S)
#: 空段落：`<w:p …/>` 与 `<w:p …></w:p>` 是同一件事的两种写法
_EMPTY_P_RE = re.compile(rb"<w:p(\s[^>]*?)?(?:/>|></w:p>)")
_RPR_RE = re.compile(rb"<w:rPr>.*?</w:rPr>", re.S)


def _blank_body(m: re.Match) -> bytes:
    """保留开闭标签与属性，把中间的**文字**换成占位符。"""
    return m.group(1) + b"\x00" + m.group(3)


def structural_fingerprint(xml: bytes) -> bytes:
    """把 XML 里**所有文字**换成占位符，只留结构。

    这就是红线三说的"文字换占位符后逐字符比对"：填值前后结构指纹必须**一模一样**。

    做三件归一化的事（两边的规则一样，所以不影响发现真变化）：

    1. `<w:t>` / `<v>` / 内联 `<t>` 的**文字**换成占位符，**属性保留**
       （属性变了也算结构变了，例如 `xml:space="preserve"`）；
    2. **丢掉"纯文字 run"**——`<w:r>` 里只有一个 `<w:t>`（外加可选的 `<w:rPr>`）。
       这种 run 是文字的容器；往空单元格里填值**必然**要新增一个，
       TA2 的口径就是"结构一致**或仅新增必要 run**"；
    3. `<w:p/>` 与 `<w:p></w:p>` 视为同一种写法（都是空段落）。
       ⚠️ 这一步必须在第 2 步**之后**：把新增的 run 丢掉以后，
       `<w:p><w:r>X</w:r></w:p>` 才变回 `<w:p></w:p>`，这时才看得出
       它和模板里那个 `<w:p/>` 是同一个空段落。
    """
    out = _PURE_TEXT_RUN_RE.sub(b"", xml)
    out = _W_T_RE.sub(_blank_body, out)
    out = _V_RE.sub(_blank_body, out)
    out = _INLINE_T_RE.sub(_blank_body, out)
    out = _EMPTY_P_RE.sub(lambda m: b"<w:p" + (m.group(1) or b"") + b"/>", out)
    return out


def count_text_runs(xml: bytes) -> int:
    """纯文字 run 的个数（填值只应让它变多，绝不该变少）。"""
    return len(_PURE_TEXT_RUN_RE.findall(xml))


def rpr_set(xml: bytes) -> set[bytes]:
    """出现过的字体/字符属性集合——用来发现"发明了新格式"。"""
    return set(_RPR_RE.findall(xml))


def package_parts(path: Path) -> dict[str, bytes]:
    with zipfile.ZipFile(path) as z:
        return {i.filename: z.read(i.filename) for i in z.infolist()}


def prove_fidelity(template: Path, written: Path) -> dict[str, Any]:
    """逐部件证明"只改了该改的文字，别的什么都没动"（红线三 / TA2）。

    四条一起成立才算通过：

    1. 包内**没有新增/删除部件**；
    2. 除 `word/document.xml` 一类含文字的部件外，其余部件**逐字节相同**；
    3. 含文字的部件：结构指纹（文字挖空 + 丢掉纯文字 run + 空段落归一）**完全一致**；
    4. 纯文字 run 的**个数没有减少**，且**没有出现模板里没有的字体属性**
       （新增的 run 必须沿用段落自己的格式，不许凭空造一个样式）。

    ⚠️ **诚实边界**：本项证明的是"排版结构没被改动"。
    它**不**逐个字节比对每个 run 的字体，也**不**渲染后比对像素——
    像素级比对要靠 `prove_format` 那套占位符回填法，见 `28 §6`。
    """
    a, b = package_parts(template), package_parts(written)
    added = sorted(set(b) - set(a))
    removed = sorted(set(a) - set(b))
    structural: list[str] = []
    text_changed: list[str] = []
    reasons: list[str] = []
    runs_added = 0
    for name in sorted(set(a) & set(b)):
        if a[name] == b[name]:
            continue
        if not (name.endswith(".xml") or name.endswith(".rels")):
            structural.append(name)  # 二进制部件不许有任何差别
            reasons.append("%s 不是文本部件却有变化" % name)
            continue
        if structural_fingerprint(a[name]) != structural_fingerprint(b[name]):
            structural.append(name)
            reasons.append("%s 的结构指纹变了" % name)
            continue
        if count_text_runs(b[name]) < count_text_runs(a[name]):
            structural.append(name)
            reasons.append("%s 少了文字 run（有内容被删掉）" % name)
            continue
        invented = rpr_set(b[name]) - rpr_set(a[name])
        if invented:
            structural.append(name)
            reasons.append("%s 出现了模板里没有的字体属性" % name)
            continue
        runs_added += count_text_runs(b[name]) - count_text_runs(a[name])
        text_changed.append(name)
    ok = not added and not removed and not structural
    note = ("结构与模板一致，只有 %s 里的文字变了（新增 %d 个文字 run，填空单元格所必需）"
            % ("、".join(text_changed), runs_added) if ok and text_changed else
            "包内没有实质差异（可能一个值都没写进去）" if ok else
            "格式/结构发生了变化，产物不可交付：" + "；".join(reasons))
    return {"ok": ok, "parts_added": added, "parts_removed": removed,
            "parts_structurally_changed": structural, "text_changed_in": text_changed,
            "text_runs_added": runs_added, "reasons": reasons, "note": note}


def opens_ok(path: Path) -> tuple[bool, str]:
    """产物能不能被办公软件打开（docx 走 python-docx，xlsx 走 openpyxl）。"""
    suffix = Path(path).suffix.lower()
    try:
        if suffix == ".docx":
            import docx

            docx.Document(str(path))
        elif suffix in (".xlsx", ".xlsm"):
            import openpyxl

            wb = openpyxl.load_workbook(str(path), read_only=True)
            wb.close()
        else:
            return True, "未做打开自检（只支持 docx/xlsx）"
    except Exception as exc:  # noqa: BLE001
        return False, "%s: %s" % (type(exc).__name__, exc)
    return True, "能打开"
