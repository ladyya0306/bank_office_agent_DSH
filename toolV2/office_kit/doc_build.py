"""Document generation and conversion: markdown/spec -> Word, PowerPoint, Excel, PDF.

Conversion strategy, in order of fidelity:
  1. Microsoft Office via COM (word/excel/ppt) — exact output, needs Office.
  2. Headless Chromium (Edge) HTML -> PDF — good typography for markdown/HTML.
  3. Pure-Python fallback — always available, reports its own lower fidelity.
"""
from __future__ import annotations

import html as _html
import json
import os
import re
import shutil
import subprocess
import tempfile
from pathlib import Path
from typing import Any

from . import office_com
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
# minimal markdown parser (headings, lists, tables, code, quotes, rules)
# --------------------------------------------------------------------------
def parse_markdown(text: str) -> list[dict[str, Any]]:
    blocks: list[dict[str, Any]] = []
    lines = text.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    i = 0
    while i < len(lines):
        raw = lines[i]
        line = raw.rstrip()
        stripped = line.strip()
        if not stripped:
            i += 1
            continue
        m = re.match(r"^(#{1,6})\s+(.*)$", stripped)
        if m:
            blocks.append({"type": "heading", "level": len(m.group(1)), "text": m.group(2).strip()})
            i += 1
            continue
        if re.match(r"^([-*_])\s*\1\s*\1[\s\-*_]*$", stripped):
            blocks.append({"type": "hr"})
            i += 1
            continue
        if stripped.startswith("```"):
            i += 1
            buf: list[str] = []
            while i < len(lines) and not lines[i].strip().startswith("```"):
                buf.append(lines[i])
                i += 1
            i += 1
            blocks.append({"type": "code", "text": "\n".join(buf)})
            continue
        if stripped.startswith(">"):
            buf = []
            while i < len(lines) and lines[i].strip().startswith(">"):
                buf.append(lines[i].strip().lstrip(">").strip())
                i += 1
            blocks.append({"type": "quote", "text": " ".join(buf)})
            continue
        if _is_table_start(lines, i):
            header = _split_row(lines[i])
            i += 2
            rows = []
            while i < len(lines) and "|" in lines[i] and lines[i].strip():
                rows.append(_split_row(lines[i]))
                i += 1
            blocks.append({"type": "table", "header": header, "data": rows})
            continue
        m = re.match(r"^\s*[-*+]\s+(.*)$", line)
        if m:
            items = []
            while i < len(lines):
                mm = re.match(r"^\s*[-*+]\s+(.*)$", lines[i])
                if not mm:
                    break
                items.append(mm.group(1).strip())
                i += 1
            blocks.append({"type": "list", "ordered": False, "items": items})
            continue
        m = re.match(r"^\s*(\d+)[.)]\s+(.*)$", line)
        if m:
            items = []
            while i < len(lines):
                mm = re.match(r"^\s*(\d+)[.)]\s+(.*)$", lines[i])
                if not mm:
                    break
                items.append(mm.group(2).strip())
                i += 1
            blocks.append({"type": "list", "ordered": True, "items": items})
            continue
        buf = [stripped]
        i += 1
        while i < len(lines) and lines[i].strip() and not _starts_block(lines, i):
            buf.append(lines[i].strip())
            i += 1
        blocks.append({"type": "paragraph", "text": " ".join(buf)})
    return blocks


def _starts_block(lines: list[str], i: int) -> bool:
    s = lines[i].strip()
    if not s:
        return True
    if re.match(r"^#{1,6}\s", s) or s.startswith("```") or s.startswith(">"):
        return True
    if re.match(r"^\s*[-*+]\s+", lines[i]) or re.match(r"^\s*\d+[.)]\s+", lines[i]):
        return True
    return _is_table_start(lines, i)


def _is_table_start(lines: list[str], i: int) -> bool:
    if i + 1 >= len(lines):
        return False
    if "|" not in lines[i]:
        return False
    sep = lines[i + 1].strip()
    return bool(re.match(r"^\|?[\s:\-|]+\|[\s:\-|]*$", sep)) and "-" in sep


def _split_row(line: str) -> list[str]:
    s = line.strip()
    if s.startswith("|"):
        s = s[1:]
    if s.endswith("|"):
        s = s[:-1]
    return [c.strip() for c in s.split("|")]


def strip_inline_md(text: str) -> str:
    t = re.sub(r"!\[([^\]]*)\]\(([^)]+)\)", r"\1", text)
    t = re.sub(r"\[([^\]]+)\]\(([^)]+)\)", r"\1", t)
    t = re.sub(r"\*\*\*(.+?)\*\*\*", r"\1", t)
    t = re.sub(r"\*\*(.+?)\*\*", r"\1", t)
    t = re.sub(r"(?<!\*)\*(?!\s)(.+?)(?<!\s)\*(?!\*)", r"\1", t)
    t = re.sub(r"`(.+?)`", r"\1", t)
    return t


# --------------------------------------------------------------------------
# inline markdown -> Word runs (keeps bold/italic)
# --------------------------------------------------------------------------
def _add_md_runs(paragraph, text: str) -> None:
    token = re.compile(r"(\*\*\*.+?\*\*\*|\*\*.+?\*\*|`[^`]+`|\*[^*\s][^*]*\*)")
    pos = 0
    for m in token.finditer(text):
        if m.start() > pos:
            paragraph.add_run(text[pos:m.start()])
        chunk = m.group(0)
        if chunk.startswith("***"):
            paragraph.add_run(chunk[3:-3]).bold = True
            paragraph.runs[-1].italic = True
        elif chunk.startswith("**"):
            paragraph.add_run(chunk[2:-2]).bold = True
        elif chunk.startswith("`"):
            r = paragraph.add_run(chunk[1:-1])
            r.font.name = "Consolas"
        else:
            paragraph.add_run(chunk[1:-1]).italic = True
        pos = m.end()
    if pos < len(text):
        paragraph.add_run(text[pos:])


# --------------------------------------------------------------------------
# Word building
# --------------------------------------------------------------------------
def _docx_set_default_font(document, font_name: str) -> None:
    """Set the document default so Chinese text uses a real CJK font, not Calibri."""
    from docx.oxml.ns import qn

    try:
        style = document.styles["Normal"]
        style.font.name = font_name
        style.element.rPr.rFonts.set(qn("w:eastAsia"), font_name)
    except Exception:  # noqa: BLE001
        pass


def build_docx_from_blocks(blocks: list[dict[str, Any]], dst: Path, *, title: str | None = None,
                           font_name: str = "微软雅黑") -> Path:
    import docx
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.shared import Pt

    document = docx.Document()
    _docx_set_default_font(document, font_name)
    if title:
        h = document.add_heading(title, level=0)
        h.alignment = WD_ALIGN_PARAGRAPH.CENTER

    for b in blocks:
        t = b.get("type")
        if t == "heading":
            document.add_heading(strip_inline_md(b["text"]), level=min(int(b.get("level", 1)), 9))
        elif t == "paragraph":
            p = document.add_paragraph()
            _add_md_runs(p, b["text"])
        elif t == "list":
            style = "List Number" if b.get("ordered") else "List Bullet"
            for item in b["items"]:
                p = document.add_paragraph(style=style)
                _add_md_runs(p, item)
        elif t == "quote":
            p = document.add_paragraph(strip_inline_md(b["text"]))
            p.style = document.styles["Quote"] if "Quote" in [s.name for s in document.styles] else p.style
            for r in p.runs:
                r.italic = True
        elif t == "code":
            p = document.add_paragraph()
            r = p.add_run(b["text"])
            r.font.name = "Consolas"
            r.font.size = Pt(9)
        elif t == "hr":
            document.add_paragraph("─" * 30)
        elif t == "pagebreak":
            document.add_page_break()
        elif t == "table":
            header = [strip_inline_md(h) for h in b.get("header", [])]
            data = b.get("data", [])
            if not header:
                continue
            table = document.add_table(rows=1, cols=len(header))
            table.style = "Light Grid Accent 1" if _has_style(document, "Light Grid Accent 1") else table.style
            for i, h in enumerate(header):
                cell = table.rows[0].cells[i]
                cell.text = ""
                run = cell.paragraphs[0].add_run(h)
                run.bold = True
            for row in data:
                cells = table.add_row().cells
                for i, v in enumerate(row[: len(header)]):
                    cells[i].text = str(v)
        elif t == "image":
            src = Path(b["path"])
            if src.exists():
                try:
                    document.add_picture(str(src), width=_emu_width(b.get("width")))
                except Exception:  # noqa: BLE001
                    document.add_paragraph(f"[图片无法插入: {src}]")
    ensure_parent(dst)
    document.save(str(dst))
    return dst


def _has_style(document, name: str) -> bool:
    try:
        return name in [s.name for s in document.styles]
    except Exception:  # noqa: BLE001
        return False


def _emu_width(spec: Any):
    from docx.shared import Cm, Inches

    if not spec:
        return Inches(6.0)
    if isinstance(spec, (int, float)):
        return Cm(float(spec))
    s = str(spec).strip().lower()
    m = re.match(r"^([\d.]+)\s*(cm|mm|in|inch|inches|pt)?$", s)
    if not m:
        return Inches(6.0)
    val = float(m.group(1))
    unit = m.group(2) or "cm"
    if unit == "cm":
        return Cm(val)
    if unit == "mm":
        return Cm(val / 10)
    if unit == "pt":
        return Inches(val / 72)
    return Inches(val)


# --------------------------------------------------------------------------
# PowerPoint building
# --------------------------------------------------------------------------
def build_pptx_from_blocks(blocks: list[dict[str, Any]], dst: Path, *, title: str | None = None,
                           bullets_per_slide: int = 6) -> Path:
    from pptx import Presentation
    from pptx.util import Inches, Pt

    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]

    if title:
        s = prs.slides.add_slide(prs.slide_layouts[0])
        s.shapes.title.text = strip_inline_md(title)
        if s.placeholders and len(s.placeholders) > 1:
            s.placeholders[1].text = ""

    current_title: str | None = None
    buf: list[tuple[int, str]] = []

    def flush():
        nonlocal buf, current_title
        if not buf:
            return
        # split long lists across several slides
        chunks = [buf[i:i + bullets_per_slide] for i in range(0, len(buf), bullets_per_slide)]
        for ci, chunk in enumerate(chunks):
            slide = prs.slides.add_slide(blank)
            _textbox(slide, current_title or (title or ""), 0.6, 0.4, 12.1, 1.1, size=26, bold=True)
            body = _textbox(slide, "", 0.8, 1.7, 11.8, 5.2, size=17)
            tf = body.text_frame
            tf.word_wrap = True
            first = True
            for lvl, txt in chunk:
                p = tf.paragraphs[0] if first else tf.add_paragraph()
                first = False
                p.text = txt
                p.level = min(lvl, 4)
                for r in p.runs:
                    r.font.size = Pt(17)
        buf = []

    for b in blocks:
        t = b.get("type")
        if t == "heading":
            flush()
            lvl = int(b.get("level", 1))
            if lvl <= 1:
                current_title = strip_inline_md(b["text"])
                buf = []
            else:
                buf.append((lvl - 2, strip_inline_md(b["text"])))
        elif t == "list":
            for item in b["items"]:
                buf.append((1, strip_inline_md(item)))
        elif t == "paragraph" and b.get("text", "").strip():
            buf.append((0, strip_inline_md(b["text"])))
        elif t == "table":
            flush()
            _pptx_table(prs, strip_inline_md(current_title or ""), b, blank)
        elif t == "pagebreak":
            flush()
    flush()

    if len(prs.slides) == 0:
        slide = prs.slides.add_slide(prs.slide_layouts[0])
        slide.shapes.title.text = title or "演示文稿"
    ensure_parent(dst)
    prs.save(str(dst))
    return dst


def _textbox(slide, text, left, top, width, height, *, size: int = 18, bold: bool = False):
    from pptx.util import Inches, Pt

    box = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    tf = box.text_frame
    tf.word_wrap = True
    tf.text = text
    for p in tf.paragraphs:
        for r in p.runs:
            r.font.size = Pt(size)
            r.font.bold = bold
    return box


def _pptx_table(prs, title_text, block, blank_layout):
    from pptx.util import Inches, Pt

    header = [strip_inline_md(h) for h in block.get("header", [])]
    data = block.get("data", [])
    if not header:
        return
    rows = min(len(data), 15) + 1
    slide = prs.slides.add_slide(blank_layout)
    if title_text:
        _textbox(slide, title_text, 0.6, 0.3, 12.1, 0.9, size=22, bold=True)
    shape = slide.shapes.add_table(rows, len(header), Inches(0.6), Inches(1.4), Inches(12.1), Inches(0.4 * rows))
    table = shape.table
    for i, h in enumerate(header):
        table.cell(0, i).text = h
    for r, row in enumerate(data[:15], start=1):
        for i, v in enumerate(row[: len(header)]):
            table.cell(r, i).text = str(v)
    for r in range(rows):
        for c in range(len(header)):
            for p in table.cell(r, c).text_frame.paragraphs:
                for run in p.runs:
                    run.font.size = Pt(11)


# --------------------------------------------------------------------------
# PDF generation from text / markdown
# --------------------------------------------------------------------------
def find_headless_browser() -> str | None:
    candidates = [
        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
    ]
    for c in candidates:
        if os.path.exists(c):
            return c
    for name in ("msedge", "chrome"):
        found = shutil.which(name)
        if found:
            return found
    return None


def html_to_pdf(html: str, dst: Path, *, browser: str | None = None, timeout: int = 90) -> Path:
    """Render HTML to PDF with headless Edge/Chrome. High quality, fonts preserved."""
    exe = browser or find_headless_browser()
    if not exe:
        raise OfficeKitError("no headless browser (Edge/Chrome) found for HTML->PDF")
    ensure_parent(dst)
    tmp_dir = Path(tempfile.mkdtemp(prefix="officekit_html_"))
    src_html = tmp_dir / "in.html"
    src_html.write_text(html, encoding="utf-8")
    out_pdf = Path(dst).resolve()
    if out_pdf.exists():
        out_pdf.unlink()
    cmd = [
        exe,
        "--headless=new",
        "--disable-gpu",
        "--no-sandbox",
        "--no-first-run",
        "--disable-extensions",
        "--virtual-time-budget=4000",
        f"--print-to-pdf={out_pdf}",
        src_html.resolve().as_uri(),
    ]
    try:
        proc = subprocess.run(cmd, capture_output=True, timeout=timeout, text=True)
    except subprocess.TimeoutExpired as exc:
        raise OfficeKitError(f"headless browser timed out after {timeout}s") from exc
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)
    if not out_pdf.exists() or out_pdf.stat().st_size == 0:
        raise OfficeKitError(
            f"headless browser did not produce a PDF (exit {proc.returncode}): {(proc.stderr or '')[:300]}"
        )
    return out_pdf


def markdown_to_html(md: str, *, title: str = "文档", font_family: str = "Microsoft YaHei",
                     css_extra: str = "") -> str:
    """Self-contained HTML with print CSS: the reliable path to a good PDF."""
    body: list[str] = []
    for b in parse_markdown(md):
        t = b.get("type")
        if t == "heading":
            lvl = min(int(b.get("level", 1)), 6)
            body.append(f"<h{lvl}>{_html.escape(strip_inline_md(b['text']))}</h{lvl}>")
        elif t == "paragraph":
            body.append(f"<p>{_html.escape(strip_inline_md(b['text']))}</p>")
        elif t == "list":
            tag = "ol" if b.get("ordered") else "ul"
            items = "".join(f"<li>{_html.escape(strip_inline_md(i))}</li>" for i in b["items"])
            body.append(f"<{tag}>{items}</{tag}>")
        elif t == "table":
            head = "".join(f"<th>{_html.escape(strip_inline_md(h))}</th>" for h in b["header"])
            rows = "".join(
                "<tr>" + "".join(f"<td>{_html.escape(str(c))}</td>" for c in r[: len(b["header"])]) + "</tr>"
                for r in b["data"]
            )
            body.append(f"<table><thead><tr>{head}</tr></thead><tbody>{rows}</tbody></table>")
        elif t == "code":
            body.append(f"<pre><code>{_html.escape(b['text'])}</code></pre>")
        elif t == "quote":
            body.append(f"<blockquote>{_html.escape(strip_inline_md(b['text']))}</blockquote>")
        elif t == "hr":
            body.append("<hr/>")
    return f"""<!DOCTYPE html>
<html lang="zh-CN"><head><meta charset="utf-8"><title>{_html.escape(title)}</title>
<style>
  @page {{ size: A4; margin: 20mm 18mm; }}
  body {{ font-family: "{font_family}", "Microsoft YaHei", "SimSun", sans-serif;
         font-size: 11pt; line-height: 1.65; color: #1a1a1a; }}
  h1 {{ font-size: 20pt; border-bottom: 2px solid #1F3864; padding-bottom: 6px; color: #1F3864; }}
  h2 {{ font-size: 15pt; color: #1F3864; margin-top: 1.3em; }}
  h3 {{ font-size: 12.5pt; color: #2E5496; }}
  table {{ border-collapse: collapse; width: 100%; margin: 10px 0; font-size: 10pt; }}
  th {{ background: #1F3864; color: #fff; border: 1px solid #BFBFBF; padding: 5px 7px; text-align: left; }}
  td {{ border: 1px solid #D9D9D9; padding: 4px 7px; vertical-align: top; }}
  tr:nth-child(even) td {{ background: #F5F8FC; }}
  pre {{ background: #F6F8FA; border: 1px solid #E1E4E8; border-radius: 4px; padding: 9px;
         font-family: Consolas, monospace; font-size: 9.5pt; white-space: pre-wrap; }}
  blockquote {{ border-left: 4px solid #1F3864; margin: 8px 0; padding: 2px 12px; color: #555; background: #F5F8FC; }}
  hr {{ border: none; border-top: 1px solid #D9D9D9; margin: 18px 0; }}
  {css_extra}
</style></head>
<body>
{chr(10).join(body)}
</body></html>"""


def build_text_pdf(text: str, dst: Path, *, title: str = "文档", font_family: str = "Microsoft YaHei") -> Path:
    """HTML -> PDF via headless browser; falls back to matplotlib text rendering."""
    md = text if re.search(r"^\s*#{1,6}\s", text, re.M) else _plain_to_md(text, title)
    html_doc = markdown_to_html(md, title=title, font_family=font_family)
    try:
        return html_to_pdf(html_doc, dst)
    except OfficeKitError:
        return _matplotlib_pdf(text, dst, title=title)


def _plain_to_md(text: str, title: str) -> str:
    lines = [f"# {title}", ""]
    for line in text.replace("\r\n", "\n").split("\n"):
        s = line.strip()
        if not s:
            lines.append("")
            continue
        if re.match(r"^第[一二三四五六七八九十百\d]+[章节部分]", s) or re.match(r"^\d+(\.\d+)*\s", s):
            lines.append(f"## {s}")
        elif s.startswith(("-", "*", "•")):
            lines.append(f"- {s.lstrip('-*• ').strip()}")
        else:
            lines.append(s)
            lines.append("")
    return "\n".join(lines)


def _matplotlib_pdf(text: str, dst: Path, *, title: str = "文档") -> Path:
    """Last-resort PDF: paginated monospaced text. Declares its own fidelity limits."""
    from matplotlib.backends.backend_pdf import PdfPages

    from .common import configure_matplotlib

    plt, font = configure_matplotlib()
    lines: list[str] = []
    for para in text.split("\n"):
        while len(para) > 46:
            lines.append(para[:46])
            para = para[46:]
        lines.append(para)

    per_page = 46
    ensure_parent(dst)
    with PdfPages(str(dst)) as pdf:
        for start in range(0, max(len(lines), 1), per_page):
            fig = plt.figure(figsize=(8.27, 11.69))
            fig.text(0.08, 0.955, title, fontsize=14, weight="bold")
            fig.text(0.08, 0.93, "─" * 60, fontsize=9)
            y = 0.90
            for line in lines[start:start + per_page]:
                fig.text(0.08, y, line, fontsize=9.5, va="top")
                y -= 0.019
            fig.text(0.08, 0.03, f"page {start // per_page + 1}", fontsize=8, color="grey")
            pdf.savefig(fig)
            plt.close(fig)
    return dst


def render_text_pdf(src: Path, dst: Path) -> Path:
    """Used as the Office-less fallback for docx->pdf."""
    from .common import read_text as _rt

    try:
        from .doc_read import docx_to_markdown, read_docx

        md = docx_to_markdown(read_docx(src))
    except Exception:  # noqa: BLE001
        md, _ = _rt(src, None)
    return build_text_pdf(md, dst, title=src.stem)


# --------------------------------------------------------------------------
# Excel building from spec
# --------------------------------------------------------------------------
def build_xlsx_from_spec(spec: dict[str, Any], dst: Path) -> Path:
    import xlsxwriter

    ensure_parent(dst)
    wb = xlsxwriter.Workbook(str(dst), {"nan_inf_to_errors": True})
    try:
        fmt_hdr = wb.add_format({"bold": True, "bg_color": "#1F3864", "font_color": "white",
                                 "border": 1, "align": "center", "valign": "vcenter"})
        fmt_title = wb.add_format({"font_size": 14, "bold": True, "font_color": "#1F3864"})
        sheets = spec.get("sheets") or [{"name": "Sheet1", "rows": spec.get("rows", [])}]
        for sh in sheets:
            name = str(sh.get("name", "Sheet1"))[:31]
            ws = wb.add_worksheet(name)
            row_i = 0
            if sh.get("title"):
                ws.write(row_i, 0, str(sh["title"]), fmt_title)
                row_i += 2
            rows = sh.get("rows") or []
            if rows and isinstance(rows[0], dict):
                header = list(rows[0].keys())
                ws.write_row(row_i, 0, header, fmt_hdr)
                for r, rec in enumerate(rows, start=1):
                    ws.write_row(row_i + r, 0, [_xl_val(rec.get(h)) for h in header])
                width_scan = [header] + [[rec.get(h) for h in header] for rec in rows[:200]]
            elif rows:
                width_scan = rows[:201]
                for r, vals in enumerate(rows):
                    vals = vals if isinstance(vals, (list, tuple)) else [vals]
                    if r == 0 and sh.get("header", True):
                        ws.write_row(row_i, 0, [_xl_val(v) for v in vals], fmt_hdr)
                    else:
                        ws.write_row(row_i + r, 0, [_xl_val(v) for v in vals])
            else:
                width_scan = []
            ws.freeze_panes(row_i + 1, 0)
            for c in range(max((len(r) if isinstance(r, (list, tuple)) else 1) for r in width_scan) if width_scan else 1):
                width = 10
                for r in width_scan[:200]:
                    if isinstance(r, (list, tuple)) and c < len(r) and r[c] is not None:
                        width = max(width, min(len(str(r[c])) * (2 if _has_cjk(str(r[c])) else 1) + 2, 50))
                ws.set_column(c, c, width)
            for chart in sh.get("charts", []) or []:
                _add_xlsx_chart(wb, ws, chart, row_i, len(width_scan))
    finally:
        wb.close()
    return dst


def _has_cjk(s: str) -> bool:
    return any(ord(c) > 0x2E80 for c in s)


def _xl_val(v):
    import math

    if v is None:
        return ""
    if isinstance(v, float) and (math.isnan(v) or math.isinf(v)):
        return ""
    if isinstance(v, (list, tuple, dict, set)):
        return json.dumps(v, ensure_ascii=False, default=str)
    return v


def _add_xlsx_chart(wb, ws, chart: dict[str, Any], header_row: int, n_rows: int) -> None:
    ctype = str(chart.get("type", "column")).lower()
    series_spec = chart.get("series") or []
    ch = wb.add_chart({"type": ctype})
    for s in series_spec:
        ch.add_series({
            "name": s.get("name", ""),
            "categories": [ws.get_name(), header_row + 1, s.get("category_col", 0),
                           header_row + max(n_rows - 1, 1), s.get("category_col", 0)],
            "values": [ws.get_name(), header_row + 1, s.get("value_col", 1),
                       header_row + max(n_rows - 1, 1), s.get("value_col", 1)],
        })
    if chart.get("title"):
        ch.set_title({"name": str(chart["title"])})
    ch.set_size({"width": chart.get("width", 640), "height": chart.get("height", 360)})
    ws.insert_chart(header_row, chart.get("at_col", 8), ch)


# --------------------------------------------------------------------------
# command: build
# --------------------------------------------------------------------------
def cmd_build(args) -> Result:
    res = Result("build")
    kind = (getattr(args, "kind", None) or "").lower()
    if kind not in ("docx", "pptx", "xlsx", "pdf", "html"):
        raise OfficeKitError("--kind must be one of docx, pptx, xlsx, pdf, html")

    spec: dict[str, Any] | None = None
    md_text: str | None = None
    if getattr(args, "spec", None):
        spec = json.loads(Path(args.spec).read_text(encoding="utf-8"))
    elif getattr(args, "markdown", None):
        md_text = Path(args.markdown).read_text(encoding="utf-8")
    elif getattr(args, "text", None):
        md_text = Path(args.text).read_text(encoding="utf-8")
    else:
        raise OfficeKitError("provide --spec (JSON) or --markdown/--text (file)")

    out = out_dir(getattr(args, "out", None), "build")
    name = safe_stem(getattr(args, "name", None) or "document")
    title = getattr(args, "title", None)
    if not title and spec:
        title = spec.get("title")
    if not title and md_text:
        m = re.search(r"^\s*#\s+(.+)$", md_text, re.M)
        title = m.group(1).strip() if m else name

    if kind == "xlsx":
        if not spec:
            raise OfficeKitError("xlsx building needs --spec")
        target = unique_path(out / f"{name}.xlsx")
        build_xlsx_from_spec(spec, target)
        res.add_artifact(target, "workbook built from spec")
        res.data.update({"kind": kind, "sheets": [s.get("name", "Sheet1") for s in spec.get("sheets", [])]})
        return res

    blocks = spec_to_blocks(spec) if spec else parse_markdown(md_text or "")
    # images referenced by spec/markdown
    if not spec and re.search(r"!\[[^\]]*\]\([^)]+\)", md_text or ""):
        blocks = _expand_md_images(blocks, md_text or "")
        res.warn("markdown images were embedded; remote URLs are not downloaded")

    if kind == "docx":
        target = unique_path(out / f"{name}.docx")
        build_docx_from_blocks(blocks, target, title=title if getattr(args, "with_title", True) else None)
        res.add_artifact(target, "Word document")
        res.data.update({"kind": kind, "blocks": len(blocks)})
        return res

    if kind == "pptx":
        target = unique_path(out / f"{name}.pptx")
        build_pptx_from_blocks(blocks, target, title=title)
        res.add_artifact(target, "PowerPoint deck")
        res.data.update({"kind": kind, "blocks": len(blocks)})
        return res

    if kind == "html":
        html_doc = markdown_to_html(md_text or blocks_to_markdown(blocks), title=title or name)
        target = unique_path(out / f"{name}.html")
        write_text(target, html_doc)
        res.add_artifact(target, "HTML document")
        res.data.update({"kind": kind})
        return res

    # pdf
    target = unique_path(out / f"{name}.pdf")
    if (getattr(args, "via", None) or "auto") == "auto":
        html_doc = markdown_to_html(md_text or blocks_to_markdown(blocks), title=title or name)
        try:
            html_to_pdf(html_doc, target)
            res.data["engine"] = "headless-browser"
        except OfficeKitError as exc:
            res.warn(f"headless HTML->PDF unavailable ({exc}); used matplotlib fallback")
            build_text_pdf(md_text or blocks_to_markdown(blocks), target, title=title or name)
            res.data["engine"] = "matplotlib-fallback"
    else:
        build_text_pdf(md_text or blocks_to_markdown(blocks), target, title=title or name)
        res.data["engine"] = "matplotlib"
    res.add_artifact(target, "PDF document")
    res.data.update({"kind": kind})
    return res


def _expand_md_images(blocks: list[dict[str, Any]], md_text: str) -> list[dict[str, Any]]:
    out: list[dict[str, Any]] = []
    for b in blocks:
        if b.get("type") == "paragraph":
            m = re.fullmatch(r"!\[([^\]]*)\]\(([^)]+)\)", b["text"].strip())
            if m and not m.group(2).lower().startswith(("http://", "https://")):
                out.append({"type": "image", "path": m.group(2), "alt": m.group(1)})
                continue
        out.append(b)
    return out


def spec_to_blocks(spec: dict[str, Any]) -> list[dict[str, Any]]:
    """Accept an explicit block list or a shorthand sections/slides structure."""
    if spec.get("blocks"):
        return spec["blocks"]
    blocks: list[dict[str, Any]] = []
    for sec in spec.get("sections", []) or []:
        if isinstance(sec, str):
            blocks.append({"type": "heading", "level": 2, "text": sec})
            continue
        if sec.get("heading"):
            blocks.append({"type": "heading", "level": sec.get("level", 2), "text": sec["heading"]})
        for para in sec.get("paragraphs", []) or []:
            blocks.append({"type": "paragraph", "text": para})
        if sec.get("bullets"):
            blocks.append({"type": "list", "ordered": False, "items": sec["bullets"]})
        for tbl in sec.get("tables", []) or []:
            blocks.append({"type": "table", "header": tbl.get("header", []), "data": tbl.get("data", [])})
        for img in sec.get("images", []) or []:
            blocks.append({"type": "image", "path": img if isinstance(img, str) else img.get("path"),
                           "width": None if isinstance(img, str) else img.get("width")})
    for slide in spec.get("slides", []) or []:
        blocks.append({"type": "heading", "level": 1, "text": slide.get("title", "")})
        if slide.get("subtitle"):
            blocks.append({"type": "paragraph", "text": slide["subtitle"]})
        if slide.get("bullets"):
            blocks.append({"type": "list", "ordered": False, "items": slide["bullets"]})
        if slide.get("notes"):
            blocks.append({"type": "quote", "text": slide["notes"]})
        for tbl in slide.get("tables", []) or []:
            blocks.append({"type": "table", "header": tbl.get("header", []), "data": tbl.get("data", [])})
    return blocks


def blocks_to_markdown(blocks: list[dict[str, Any]]) -> str:
    L: list[str] = []
    for b in blocks:
        t = b.get("type")
        if t == "heading":
            L += ["#" * min(int(b.get("level", 1)), 6) + " " + b["text"], ""]
        elif t == "paragraph":
            L += [b["text"], ""]
        elif t == "list":
            L += [("- " if not b.get("ordered") else "1. ") + i for i in b["items"]]
            L.append("")
        elif t == "table":
            header = b.get("header", [])
            L.append("| " + " | ".join(str(h) for h in header) + " |")
            L.append("| " + " | ".join("---" for _ in header) + " |")
            for r in b.get("data", []):
                L.append("| " + " | ".join(str(c) for c in r[: len(header)]) + " |")
            L.append("")
        elif t == "quote":
            L += [f"> {b['text']}", ""]
        elif t == "code":
            L += ["```", b["text"], "```", ""]
        elif t == "image":
            L += [f"![{b.get('alt', '')}]({b['path']})", ""]
        elif t == "hr":
            L += ["---", ""]
        elif t == "pagebreak":
            L += ["", "---", ""]
    return "\n".join(L)


# --------------------------------------------------------------------------
# command: convert
# --------------------------------------------------------------------------
CONVERT_PAIRS = {
    ("docx", "pdf"), ("doc", "pdf"), ("xlsx", "pdf"), ("xls", "pdf"), ("pptx", "pdf"), ("ppt", "pdf"),
    ("docx", "docx"), ("doc", "docx"), ("doc", "docx"), ("xls", "xlsx"), ("ppt", "pptx"),
    ("docx", "html"), ("doc", "html"), ("md", "docx"), ("txt", "docx"),
    ("md", "pdf"), ("txt", "pdf"), ("html", "pdf"), ("md", "pptx"), ("md", "html"),
    ("xlsx", "csv"), ("xls", "csv"), ("pptx", "png"),
}


def cmd_convert(args) -> Result:
    res = Result("convert")
    files = resolve_inputs(args.input)
    to = (getattr(args, "to", None) or "").lower().lstrip(".")
    if not to:
        raise OfficeKitError("--to is required (pdf, docx, xlsx, pptx, csv, html, md, png)")
    out = out_dir(getattr(args, "out", None), "convert")
    mode = (getattr(args, "engine", None) or "auto").lower()
    res.data["office_available"] = office_com.availability_report()
    results: list[dict[str, Any]] = []

    for f in files:
        src_fmt = f.suffix.lower().lstrip(".")
        entry: dict[str, Any] = {"source": str(f.resolve()), "from": src_fmt, "to": to}
        target = unique_path(out / f"{safe_stem(f.stem)}.{to}")
        try:
            engine_used = _convert_one(f, target, src_fmt, to, mode, res)
            entry.update({"ok": True, "target": str(target), "engine": engine_used})
            res.add_artifact(target, f"{f.name} -> {to} ({engine_used})")
        except OfficeKitError as exc:
            entry.update({"ok": False, "error": str(exc)})
            res.warn(f"{f.name}: {exc}")
        results.append(entry)

    res.data["converted"] = results
    res.data["succeeded"] = sum(1 for r in results if r.get("ok"))
    res.data["failed"] = sum(1 for r in results if not r.get("ok"))
    return res


def _convert_one(src: Path, target: Path, src_fmt: str, to: str, mode: str, res: Result) -> str:
    # ---- markdown / text / html sources: pure python builds
    if src_fmt in ("md", "markdown", "txt", "text", "html", "htm"):
        content = src.read_text(encoding="utf-8", errors="replace")
        if src_fmt in ("html", "htm"):
            html_doc = content
            md = content
        else:
            md = content
            html_doc = markdown_to_html(content, title=src.stem)
        if to == "docx":
            blocks = parse_markdown(md) if src_fmt != "html" else [{"type": "paragraph", "text": re.sub(r"<[^>]+>", "", content)}]
            build_docx_from_blocks(blocks, target, title=None)
            return "python-docx"
        if to == "pptx":
            build_pptx_from_blocks(parse_markdown(md), target, title=src.stem)
            return "python-pptx"
        if to == "pdf":
            try:
                html_to_pdf(html_doc, target)
                return "headless-browser"
            except OfficeKitError:
                build_text_pdf(md, target, title=src.stem)
                return "matplotlib-fallback"
        if to in ("html", "htm"):
            write_text(target, html_doc)
            return "python"
        raise OfficeKitError(f"unsupported conversion {src_fmt} -> {to}")

    # ---- Office document sources
    if src_fmt in ("doc", "docx") and to == "pdf":
        if mode in ("auto", "office") and office_com.available("word"):
            office_com.word_to_pdf(src, target)
            return "ms-word"
        if mode == "office":
            raise OfficeKitError("MS Word unavailable and engine=office was required")
        res.warn(f"{src.name}: Word unavailable, used lower-fidelity text-layout PDF")
        render_text_pdf(src, target)
        return "python-fallback"

    if src_fmt in ("doc", "docx") and to in ("docx", "html", "txt", "rtf"):
        fmt = {"docx": 16, "html": 8, "txt": 2, "rtf": 6}[to]
        if office_com.available("word"):
            office_com.word_save_as(src, target, fmt)
            return "ms-word"
        if to == "docx":
            raise OfficeKitError("converting legacy .doc needs Microsoft Word")
        if to in ("html", "txt"):
            from .doc_read import docx_to_markdown, read_docx

            md = docx_to_markdown(read_docx(src))
            write_text(target, md if to == "html" else re.sub(r"[#*`]", "", md))
            return "python-fallback"
        raise OfficeKitError(f"cannot convert {src_fmt} -> {to} without Word")

    if src_fmt in ("xlsx", "xls", "xlsm") and to == "pdf":
        if mode in ("auto", "office") and office_com.available("excel"):
            office_com.excel_to_pdf(src, target)
            return "ms-excel"
        raise OfficeKitError("xlsx -> pdf requires Microsoft Excel")

    if src_fmt in ("xlsx", "xls", "xlsm") and to == "csv":
        from . import data_ops

        df = data_ops.read_table(src, header=0)
        df.to_csv(target, index=False, encoding="utf-8-sig")
        return "pandas"
    if src_fmt in ("xls",) and to == "xlsx":
        if office_com.available("excel"):
            office_com.excel_to_format(src, target, office_com.XL_XLSX)
            return "ms-excel"
        import pandas as pd

        pd.read_excel(src, sheet_name=None).items()
        with __import__("pandas").ExcelWriter(target, engine="xlsxwriter") as xw:
            for name, df in pd.read_excel(src, sheet_name=None).items():
                df.to_excel(xw, sheet_name=name[:31], index=False)
        return "pandas"

    if src_fmt in ("pptx", "ppt") and to == "pdf":
        if mode in ("auto", "office") and office_com.available("ppt"):
            office_com.ppt_to_pdf(src, target)
            return "ms-powerpoint"
        raise OfficeKitError("pptx -> pdf requires Microsoft PowerPoint")
    if src_fmt in ("pptx", "ppt") and to == "png":
        if office_com.available("ppt"):
            imgs = office_com.ppt_to_images(src, target.parent / f"{safe_stem(src.stem)}_slides")
            if imgs:
                shutil.copy2(imgs[0], target)
                return "ms-powerpoint"
        raise OfficeKitError("pptx -> png requires Microsoft PowerPoint")
    if src_fmt == "ppt" and to == "pptx":
        if not office_com.available("ppt"):
            raise OfficeKitError("ppt -> pptx requires Microsoft PowerPoint")
        office_com.ppt_to_pdf(src, target.with_suffix(".pdf"))
        raise OfficeKitError("PowerPoint cannot SaveAs pptx via this path; convert to PDF instead")

    if src_fmt == "pdf" and to in ("docx", "xlsx", "txt", "md"):
        return _pdf_to_office(src, target, to, res)

    raise OfficeKitError(
        f"unsupported conversion {src_fmt} -> {to}; "
        f"supported: {sorted({f'{a}->{b}' for a, b in CONVERT_PAIRS})}"
    )


def _pdf_to_office(src: Path, target: Path, to: str, res: Result) -> str:
    """PDF -> docx/xlsx/txt/md by extracting text (and OCR when the PDF is a scan)."""
    from .doc_read import read_pdf_text, read_pdf_tables

    pages, meta = read_pdf_text(src)
    needs_ocr = meta.get("looks_scanned")
    if needs_ocr:
        res.warn(f"{src.name}: little extractable text; running OCR to recover content")
        try:
            from . import ocr_ops

            pages = ocr_ops.ocr_pdf_pages(src)
            meta["ocr_used"] = True
        except OfficeKitError as exc:
            res.warn(f"OCR failed ({exc}); output contains only the extractable text")

    if to in ("txt", "md"):
        text = "\n\n".join(p.get("text", "") for p in pages)
        if to == "md":
            text = f"# {src.stem}\n\n" + text
        write_text(target, text)
        return "python+ocr" if needs_ocr else "python"
    if to == "docx":
        blocks: list[dict[str, Any]] = []
        for p in pages:
            blocks.append({"type": "heading", "level": 2, "text": f"第 {p['page']} 页"})
            for para in _split_paragraphs(p.get("text", "")):
                blocks.append({"type": "paragraph", "text": para})
        build_docx_from_blocks(blocks, target, title=src.stem)
        return "python+ocr" if needs_ocr else "python"
    if to == "xlsx":
        import xlsxwriter

        tables = read_pdf_tables(src)
        pairs: list[tuple[str, list[list[str]]]] = []
        if tables:
            for i, t in enumerate(tables, start=1):
                pairs.append((f"table_{i}", [t["header"]] + t["data"]))
        else:
            for p in pages:
                rows = [ln.split() for ln in p.get("text", "").splitlines() if ln.strip()]
                if rows:
                    pairs.append((f"page_{p['page']}", rows))
            res.warn("no ruled tables detected; each page was placed as whitespace-split columns")
        with xlsxwriter.Workbook(str(target)) as wb:
            for sheet, rows in pairs:
                ws = wb.add_worksheet(sheet[:31])
                for r, row in enumerate(rows):
                    ws.write_row(r, 0, [_xl_val(c) for c in row])
        return "python+ocr" if needs_ocr else "python"
    raise OfficeKitError(f"pdf -> {to} is not supported")


def _split_paragraphs(text: str) -> list[str]:
    paras = [p.strip() for p in re.split(r"\n\s*\n", text) if p.strip()]
    if len(paras) <= 1:
        paras = [ln.strip() for ln in text.splitlines() if ln.strip()]
    return paras


# --------------------------------------------------------------------------
# command: pdf operations
# --------------------------------------------------------------------------
def cmd_pdf(args) -> Result:
    res = Result("pdf")
    op = (getattr(args, "op", None) or "").lower()
    files = resolve_inputs(args.input)
    out = out_dir(getattr(args, "out", None), "pdf")
    password = getattr(args, "password", None)
    res.data["op"] = op

    if op == "merge":
        import pypdf

        writer = pypdf.PdfWriter()
        for f in files:
            reader = pypdf.PdfReader(str(f))
            if reader.is_encrypted:
                reader.decrypt(password or "")
            for page in reader.pages:
                writer.add_page(page)
        target = unique_path(out / f"{safe_stem(getattr(args, 'name', None) or 'merged')}.pdf")
        ensure_parent(target)
        with open(target, "wb") as fh:
            writer.write(fh)
        res.add_artifact(target, f"merged {len(files)} PDFs")
        res.data.update({"sources": [str(f) for f in files], "pages": len(writer.pages)})
        return res

    if op == "split":
        import pypdf

        src = files[0]
        reader = pypdf.PdfReader(str(src))
        if reader.is_encrypted:
            reader.decrypt(password or "")
        every = int(getattr(args, "every", 1) or 1)
        made: list[str] = []
        n = len(reader.pages)
        for start in range(0, n, every):
            writer = pypdf.PdfWriter()
            for p in range(start, min(start + every, n)):
                writer.add_page(reader.pages[p])
            target = unique_path(out / f"{safe_stem(src.stem)}_p{start + 1}-{min(start + every, n)}.pdf")
            ensure_parent(target)
            with open(target, "wb") as fh:
                writer.write(fh)
            res.add_artifact(target, f"pages {start + 1}-{min(start + every, n)}")
            made.append(str(target))
        res.data.update({"source": str(src), "total_pages": n, "chunks": len(made)})
        return res

    if op in ("extract-pages", "rotate", "encrypt", "decrypt", "info", "images"):
        import pypdf

        src = files[0]
        reader = pypdf.PdfReader(str(src))
        if reader.is_encrypted:
            if not reader.decrypt(password or ""):
                raise OfficeKitError("PDF is encrypted; pass --password")
        if op == "info":
            from .doc_read import _pdf_meta

            meta = _pdf_meta(reader)
            res.data["info"] = meta
            res.data["page_sizes"] = [
                {"page": i, "width_pt": round(float(p.mediabox.width), 1),
                 "height_pt": round(float(p.mediabox.height), 1)}
                for i, p in enumerate(reader.pages[:200], start=1)
            ]
            if meta.get("looks_scanned"):
                res.warn("this PDF looks like a scan; use `ocr pdf` to get text")
            return res

        if op == "images":
            from .doc_read import extract_pdf_images

            found = extract_pdf_images(src, out / "images", password=password)
            res.data["images"] = found
            if not found:
                res.warn("no embedded raster images found (vector pages are not extracted here)")
            else:
                res.add_artifact(out / "images", f"{len(found)} extracted images")
            return res

        writer = pypdf.PdfWriter()
        pages = _parse_page_ranges(getattr(args, "pages", None), len(reader.pages))
        for p in pages:
            writer.add_page(reader.pages[p - 1])
        if op == "rotate":
            angle = int(getattr(args, "angle", 90) or 90)
            for page in writer.pages:
                page.rotate(angle)
        if op == "encrypt":
            if not getattr(args, "new_password", None):
                raise OfficeKitError("encrypt requires --new-password")
            writer.encrypt(args.new_password)
        suffix = {"extract-pages": "pages", "rotate": f"rot{getattr(args, 'angle', 90)}", "encrypt": "encrypted"}[op]
        target = unique_path(out / f"{safe_stem(src.stem)}_{suffix}.pdf")
        ensure_parent(target)
        with open(target, "wb") as fh:
            writer.write(fh)
        res.add_artifact(target, f"{op} result")
        res.data.update({"source": str(src), "pages": pages})
        return res

    if op == "to-images":
        from .doc_read import render_pdf_pages

        src = files[0]
        pages = _parse_page_ranges(getattr(args, "pages", None), None)
        found = render_pdf_pages(src, out / "images", pages=pages or None,
                                 scale=float(getattr(args, "scale", 2.0) or 2.0),
                                 fmt=getattr(args, "image_format", "png"), password=password)
        res.data["images"] = found
        res.add_artifact(out / "images", f"{len(found)} rendered pages")
        return res

    raise OfficeKitError("--op must be one of merge, split, extract-pages, rotate, encrypt, info, images, to-images")


def _parse_page_ranges(spec: str | None, total: int | None) -> list[int]:
    """'1-3,5,8-10' -> [1,2,3,5,8,9,10]."""
    if not spec:
        if total is None:
            return []
        return list(range(1, total + 1))
    out: list[int] = []
    for part in str(spec).split(","):
        part = part.strip()
        if not part:
            continue
        m = re.match(r"^(\d+)\s*-\s*(\d+)$", part)
        if m:
            a, b = int(m.group(1)), int(m.group(2))
            out.extend(range(min(a, b), max(a, b) + 1))
        elif part.isdigit():
            out.append(int(part))
        else:
            raise OfficeKitError(f"bad page range: {part!r}")
    if total:
        out = [p for p in out if 1 <= p <= total]
    return sorted(set(out))


# --------------------------------------------------------------------------
# command: template fill
# --------------------------------------------------------------------------
def cmd_fill(args) -> Result:
    res = Result("fill")
    template = Path(getattr(args, "template", None) or "")
    if not template.exists():
        raise OfficeKitError(f"template not found: {template}")
    if not getattr(args, "data", None):
        raise OfficeKitError("--data (JSON file) is required")

    payload = json.loads(Path(args.data).read_text(encoding="utf-8"))
    if getattr(args, "repeat", None):
        rows = payload if isinstance(payload, list) else payload.get("rows", [])
        if not isinstance(rows, list):
            raise OfficeKitError("--repeat expects a list of records")
    else:
        rows = [payload] if isinstance(payload, dict) else [{"items": payload}]

    out = out_dir(getattr(args, "out", None), "fill")
    stem = safe_stem(getattr(args, "name", None) or template.stem)
    suffix = template.suffix.lower()
    produced: list[str] = []

    if suffix == ".docx":
        try:
            import docxtpl
        except ImportError as exc:
            raise OfficeKitError(
                "docxtpl is unavailable (needs pkg_resources from setuptools<81); "
                f"run: python -m pip install \"setuptools<81\" docxtpl  [{exc}]"
            ) from exc
        for i, row in enumerate(rows, start=1):
            ctx = dict(row) if isinstance(row, dict) else {"value": row}
            doc = docxtpl.DocxTemplate(str(template))
            try:
                doc.render(ctx)
            except Exception as exc:  # noqa: BLE001
                raise OfficeKitError(
                    f"template render failed on record {i}: {exc}. "
                    f"Available keys: {sorted(ctx)}"
                ) from exc
            name = f"{stem}_{i}.docx" if len(rows) > 1 else f"{stem}_filled.docx"
            target = unique_path(out / name)
            ensure_parent(target)
            doc.save(str(target))
            res.add_artifact(target, f"filled record {i}")
            produced.append(str(target))
        res.data.update({"template": str(template), "records": len(rows), "outputs": len(produced)})
        return res

    if suffix == ".xlsx":
        import openpyxl

        wb = openpyxl.load_workbook(str(template))
        count = 0
        for ws in wb.worksheets:
            for row in ws.iter_rows():
                for cell in row:
                    if isinstance(cell.value, str) and "{{" in cell.value and "}}" in cell.value:
                        cell.value = _render_inline(cell.value, rows[0] if rows else {})
                        count += 1
        target = unique_path(out / f"{stem}_filled.xlsx")
        ensure_parent(target)
        wb.save(str(target))
        res.add_artifact(target, f"filled workbook ({count} placeholders)")
        res.data.update({"template": str(template), "placeholders_filled": count})
        if count == 0:
            res.warn("no {{placeholder}} cells were found in the workbook")
        return res

    if suffix == ".pptx":
        from pptx import Presentation

        prs = Presentation(str(template))
        ctx = rows[0] if rows else {}
        count = 0
        for slide in prs.slides:
            for shape in slide.shapes:
                if shape.has_text_frame:
                    for para in shape.text_frame.paragraphs:
                        for run in para.runs:
                            if "{{" in run.text:
                                run.text = _render_inline(run.text, ctx)
                                count += 1
        target = unique_path(out / f"{stem}_filled.pptx")
        ensure_parent(target)
        prs.save(str(target))
        res.add_artifact(target, f"filled deck ({count} placeholders)")
        res.data.update({"template": str(template), "placeholders_filled": count})
        return res

    raise OfficeKitError(f"template filling supports .docx/.xlsx/.pptx, got {suffix}")


def _render_inline(text: str, ctx: dict[str, Any]) -> str:
    def repl(m):
        key = m.group(1).strip()
        cur: Any = ctx
        for part in re.split(r"[.\[\]]+", key):
            if not part:
                continue
            if isinstance(cur, dict) and part in cur:
                cur = cur[part]
            else:
                return m.group(0)
        return str(cur)

    return re.sub(r"\{\{\s*([^}]+?)\s*\}\}", repl, text)


# --------------------------------------------------------------------------
# command: to-docx  (legacy .doc/.xls/.ppt -> modern, via Office)
# --------------------------------------------------------------------------
LEGACY_TARGETS = {".doc": ".docx", ".xls": ".xlsx", ".ppt": ".pptx"}
FILE_FORMATS = {".docx": 16, ".xlsx": 51, ".pptx": 24}


def cmd_to_docx(args) -> Result:
    """Convert legacy Office formats using one reused application instance.

    Pure-Python readers cannot open the binary .doc/.xls/.ppt container, so this
    is a required preprocessing step before format-preserving editing.
    """
    res = Result("to-docx")
    files = resolve_inputs(args.input)
    out = out_dir(getattr(args, "out", None), "to-docx")
    res.data["office"] = office_com.availability_report()
    results: list[dict[str, Any]] = []

    legacy = [f for f in files if f.suffix.lower() in LEGACY_TARGETS]
    already = [f for f in files if f.suffix.lower() in (".docx", ".xlsx", ".pptx")]
    unsupported = [f for f in files if f not in legacy and f not in already]
    for f in unsupported:
        res.warn(f"{f.name}: not an Office file, skipped")

    groups: dict[str, list[tuple[Path, Path]]] = {"word": [], "excel": [], "ppt": []}
    for f in legacy:
        ext = f.suffix.lower()
        target_ext = LEGACY_TARGETS[ext]
        dst = unique_path(out / f"{safe_stem(f.stem)}{target_ext}")
        if ext == ".doc":
            groups["word"].append((f, dst))
        elif ext == ".xls":
            groups["excel"].append((f, dst))
        else:
            groups["ppt"].append((f, dst))

    for f in already:
        # Nothing to convert; copy so the output directory is a complete set.
        dst = unique_path(out / f.name)
        ensure_parent(dst)
        shutil.copy2(f, dst)
        results.append({"source": str(f), "target": str(dst), "ok": True,
                        "note": "already modern format; copied"})
        res.add_artifact(dst, f"{f.name}（已是新格式，直接复制）")

    try:
        if groups["word"]:
            if not office_com.available("word"):
                raise OfficeKitError("converting .doc needs Microsoft Word (COM unavailable)")
            _convert_batch("word", groups["word"], results, res)
        if groups["excel"]:
            if not office_com.available("excel"):
                raise OfficeKitError("converting .xls needs Microsoft Excel")
            _convert_batch("excel", groups["excel"], results, res)
        if groups["ppt"]:
            if not office_com.available("ppt"):
                raise OfficeKitError("converting .ppt needs Microsoft PowerPoint")
            _convert_batch("ppt", groups["ppt"], results, res)
    except OfficeKitError as exc:
        res.warn(str(exc))

    res.data["converted"] = results
    res.data["succeeded"] = sum(1 for r in results if r.get("ok"))
    res.data["failed"] = sum(1 for r in results if not r.get("ok"))
    return res


def _convert_batch(app: str, pairs: list[tuple[Path, Path]], results, res: Result) -> None:
    """Save-As each legacy file to its modern format, reusing one app instance."""
    from . import office_com as com

    openers = {
        "word": lambda application, p: application.Documents.Open(
            str(p.resolve()), ReadOnly=True, AddToRecentFiles=False),
        "excel": lambda application, p: application.Workbooks.Open(
            str(p.resolve()), ReadOnly=True, UpdateLinks=0),
        "ppt": lambda application, p: application.Presentations.Open(
            str(p.resolve()), WithWindow=0, ReadOnly=-1),
    }
    savers = {
        "word": lambda doc, p, fmt: doc.SaveAs2(str(p.resolve()), FileFormat=fmt),
        "excel": lambda doc, p, fmt: doc.SaveAs(str(p.resolve()), FileFormat=fmt),
        "ppt": lambda doc, p, fmt: doc.SaveAs(str(p.resolve()), fmt),
    }

    with com.office_app(app) as application:
        for src, dst in pairs:
            entry: dict[str, Any] = {"source": str(src), "target": str(dst)}
            doc = None
            try:
                ensure_parent(dst)
                fmt = FILE_FORMATS[dst.suffix.lower()]
                doc = openers[app](application, src)
                savers[app](doc, dst, fmt)
                entry["ok"] = True
                res.add_artifact(dst, f"{src.name} → {dst.name}")
            except Exception as exc:  # noqa: BLE001
                entry["ok"] = False
                entry["error"] = str(exc)
                res.warn(f"{src.name}: 转换失败 {exc}")
            finally:
                # Release the document reference before the app quits, otherwise
                # a zombie WINWORD.EXE/EXCEL.EXE keeps the file locked.
                if doc is not None:
                    import contextlib as _ctx

                    with _ctx.suppress(Exception):
                        doc.Close(False)
                    del doc
                    com._release_com()
            results.append(entry)
