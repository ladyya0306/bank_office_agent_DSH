"""Document parsing: PDF / Word / PowerPoint / Excel -> text, markdown, tables, JSON.

Reading is done with pure-Python libraries so it always works; the parts where
pure Python genuinely cannot compete (exact pagination, exotic embedded
objects) are reported as explicit warnings rather than guessed at.
"""
from __future__ import annotations

import json
import re
import zipfile
from pathlib import Path
from typing import Any

from .common import OfficeKitError, Result, out_dir, read_table, resolve_inputs, safe_stem, unique_path, write_text

# --------------------------------------------------------------------------
# PDF
# --------------------------------------------------------------------------
def _pdf_meta(reader) -> dict[str, Any]:
    meta: dict[str, Any] = {"pages": len(reader.pages), "encrypted": bool(reader.is_encrypted)}
    info = {}
    try:
        info = dict(reader.metadata or {})
    except Exception:  # noqa: BLE001
        pass
    for k, v in info.items():
        key = str(k).lstrip("/")
        try:
            meta[key] = str(v)
        except Exception:  # noqa: BLE001
            meta[key] = repr(v)
    return meta


def read_pdf_text(path: Path, *, password: str | None = None) -> tuple[list[dict[str, Any]], dict[str, Any]]:
    """Extract per-page text and metadata. Returns (pages, meta)."""
    import pypdf

    reader = pypdf.PdfReader(str(path))
    if reader.is_encrypted:
        try:
            ok = reader.decrypt(password or "")
        except Exception as exc:  # noqa: BLE001
            raise OfficeKitError(f"PDF is encrypted and could not be opened: {exc}") from exc
        if not ok:
            raise OfficeKitError("PDF is encrypted; pass --password")
    meta = _pdf_meta(reader)
    pages: list[dict[str, Any]] = []
    total_chars = 0
    for i, page in enumerate(reader.pages, start=1):
        try:
            text = page.extract_text() or ""
        except Exception as exc:  # noqa: BLE001
            text = ""
            meta.setdefault("page_errors", []).append({"page": i, "error": str(exc)})
        total_chars += len(text.strip())
        pages.append(
            {
                "page": i,
                "text": text,
                "chars": len(text.strip()),
                "images": _count_page_images(page),
                "rotation": int(page.get("/Rotate", 0) or 0),
            }
        )
    meta["total_chars"] = total_chars
    # A PDF where no page yields a meaningful amount of text is a scan.
    meta["looks_scanned"] = total_chars < max(50, 20 * len(pages))
    return pages, meta


def _count_page_images(page) -> int:
    try:
        return len(list(page.images))
    except Exception:  # noqa: BLE001
        return 0


def read_pdf_tables(path: Path) -> list[dict[str, Any]]:
    """Extract ruled/positional tables. Prefers pdfplumber when present."""
    try:
        import pdfplumber  # type: ignore
    except ImportError:
        return []
    out: list[dict[str, Any]] = []
    try:
        with pdfplumber.open(str(path)) as pdf:
            for pno, page in enumerate(pdf.pages, start=1):
                for tno, table in enumerate(page.extract_tables() or [], start=1):
                    rows = [[("" if c is None else str(c).strip()) for c in row] for row in table]
                    rows = [r for r in rows if any(c for c in r)]
                    if len(rows) < 2:
                        continue
                    header = rows[0]
                    out.append(
                        {
                            "page": pno,
                            "table": tno,
                            "rows": len(rows) - 1,
                            "columns": len(header),
                            "header": header,
                            "data": rows[1:],
                        }
                    )
    except Exception as exc:  # noqa: BLE001
        raise OfficeKitError(f"table extraction failed: {exc}") from exc
    return out


def extract_pdf_images(path: Path, out: Path, *, min_bytes: int = 4096, password: str | None = None) -> list[dict[str, Any]]:
    """Pull embedded raster images out of a PDF."""
    import pypdf

    out.mkdir(parents=True, exist_ok=True)
    reader = pypdf.PdfReader(str(path))
    if reader.is_encrypted:
        reader.decrypt(password or "")
    found: list[dict[str, Any]] = []
    for pno, page in enumerate(reader.pages, start=1):
        try:
            images = list(page.images)
        except Exception:  # noqa: BLE001
            continue
        for ino, img in enumerate(images, start=1):
            data = img.data
            if len(data) < min_bytes:
                continue
            name = Path(getattr(img, "name", f"p{pno}_i{ino}.png")).name
            target = unique_path(out / f"page{pno:03d}_{ino:02d}_{safe_stem(Path(name).stem)}.png")
            target.write_bytes(data)
            found.append({"page": pno, "file": str(target), "bytes": len(data)})
    return found


def render_pdf_pages(path: Path, out: Path, *, pages: list[int] | None = None, scale: float = 2.0,
                     fmt: str = "png", password: str | None = None) -> list[dict[str, Any]]:
    """Rasterize PDF pages. pdfium is closed explicitly: it locks the file on Windows."""
    import pypdfium2 as pdfium

    from .common import ensure_parent

    out.mkdir(parents=True, exist_ok=True)
    doc = pdfium.PdfDocument(str(path), password=password) if password else pdfium.PdfDocument(str(path))
    produced: list[dict[str, Any]] = []
    try:
        total = len(doc)
        wanted = pages if pages else list(range(1, total + 1))
        for pno in wanted:
            if pno < 1 or pno > total:
                continue
            page = doc[pno - 1]
            try:
                bitmap = page.render(scale=scale)
                image = bitmap.to_pil()
                try:
                    target = out / f"page_{pno:04d}.{fmt}"
                    ensure_parent(target)
                    image.save(target)
                    produced.append({"page": pno, "file": str(target), "size": list(image.size)})
                finally:
                    with_close = getattr(image, "close", None)
                    if with_close:
                        with_close()
            finally:
                try:
                    page.close()
                except Exception:  # noqa: BLE001
                    pass
    finally:
        doc.close()
    return produced


def pdf_word_count(pages: list[dict[str, Any]]) -> dict[str, int]:
    text = "\n".join(p["text"] for p in pages)
    cjk = len(re.findall(r"[\u4e00-\u9fff]", text))
    latin = len(re.findall(r"[A-Za-z]+", text))
    return {"cjk_characters": cjk, "latin_words": latin, "total_chars": len(text)}


# --------------------------------------------------------------------------
# Word
# --------------------------------------------------------------------------
def read_docx(path: Path) -> dict[str, Any]:
    import docx

    document = docx.Document(str(path))
    blocks: list[dict[str, Any]] = []

    # Body order matters: iterate the XML so tables and paragraphs interleave.
    body = document.element.body
    para_map = {p._element: p for p in document.paragraphs}
    table_map = {t._element: t for t in document.tables}

    for child in body.iterchildren():
        if child in para_map:
            p = para_map[child]
            blocks.append(_docx_paragraph(p))
        elif child in table_map:
            t = table_map[child]
            blocks.append(_docx_table(t))

    headings = [b for b in blocks if b["type"] == "heading"]
    tables = [b for b in blocks if b["type"] == "table"]
    core = document.core_properties
    return {
        "paragraphs": sum(1 for b in blocks if b["type"] in ("heading", "paragraph", "list_item")),
        "headings": headings,
        "tables": tables,
        "blocks": blocks,
        "characters": sum(len(b.get("text", "")) for b in blocks),
        "inline_images": count_docx_images(document),
        "metadata": {
            "title": core.title or "",
            "author": core.author or "",
            "created": str(core.created or ""),
            "modified": str(core.modified or ""),
            "last_modified_by": core.last_modified_by or "",
            "revision": core.revision,
        },
        "sections": [
            {
                "page_width_cm": round(s.page_width.cm, 2) if s.page_width else None,
                "page_height_cm": round(s.page_height.cm, 2) if s.page_height else None,
                "orientation": "landscape" if s.orientation == 1 else "portrait",
            }
            for s in document.sections
        ],
        "headers_footers": _docx_headers_footers(document),
    }


def count_docx_images(document) -> int:
    count = 0
    for rel in document.part.rels.values():
        if "image" in rel.reltype:
            count += 1
    return count


def _docx_headers_footers(document) -> dict[str, list[str]]:
    out: dict[str, list[str]] = {"header": [], "footer": []}
    for s in document.sections:
        for label, part in (("header", s.header), ("footer", s.footer)):
            try:
                txt = "\n".join(p.text for p in part.paragraphs if p.text.strip())
                if txt.strip():
                    out[label].append(txt.strip())
            except Exception:  # noqa: BLE001
                continue
    return out


def _docx_paragraph(p) -> dict[str, Any]:
    style = (p.style.name if p.style is not None else "") or ""
    level = None
    m = re.match(r"Heading (\d+)", style)
    if m:
        level = int(m.group(1))
    elif style.lower() in ("title", "标题"):
        level = 1
    elif style.lower().startswith("heading"):
        level = 1
    else:
        # Chinese Word builds report 标题 1 / 标题 2
        m2 = re.match(r"标题\s*(\d+)", style)
        if m2:
            level = int(m2.group(1))

    is_list = _is_list_paragraph(p) or style.lower().startswith(("list", "bullet", "列表"))
    text = p.text
    kind = "heading" if level else ("list_item" if is_list else "paragraph")
    entry: dict[str, Any] = {"type": kind, "text": text, "style": style}
    if level:
        entry["level"] = level
    if kind == "paragraph" and not text.strip():
        entry["type"] = "blank"
    try:
        if p._element.pPr is not None:
            numpr = p._element.pPr.numPr
            if numpr is not None and numpr.numId is not None:
                entry.setdefault("numbering_id", int(numpr.numId.val))
    except Exception:  # noqa: BLE001
        pass
    runs = [r.text for r in p.runs if r.text]
    if runs and "".join(runs) != text:
        entry["run_texts"] = runs
    return entry


def _is_list_paragraph(p) -> bool:
    try:
        if p._element.pPr is not None and p._element.pPr.numPr is not None:
            return True
    except Exception:  # noqa: BLE001
        pass
    style = (p.style.name if p.style is not None else "") or ""
    return style.lower() in ("list paragraph", "list bullet", "list number")


def _docx_table(t) -> dict[str, Any]:
    rows: list[list[str]] = []
    for row in t.rows:
        rows.append([c.text.strip() for c in row.cells])
    header = rows[0] if rows else []
    return {
        "type": "table",
        "rows": len(rows),
        "columns": len(header),
        "header": header,
        "data": rows[1:] if rows else [],
        "style": (t.style.name if t.style is not None else "") or "",
    }


# --------------------------------------------------------------------------
# PowerPoint
# --------------------------------------------------------------------------
def read_pptx(path: Path) -> dict[str, Any]:
    from pptx import Presentation

    prs = Presentation(str(path))
    slides: list[dict[str, Any]] = []
    for i, slide in enumerate(prs.slides, start=1):
        texts: list[str] = []
        tables: list[dict[str, Any]] = []
        picture_count = 0
        for shape in slide.shapes:
            if shape.has_text_frame and shape.text_frame.text.strip():
                texts.append(shape.text_frame.text.strip())
            if getattr(shape, "has_table", False) and shape.has_table:
                rows = [[c.text.strip() for c in r.cells] for r in shape.table.rows]
                tables.append({"header": rows[0] if rows else [], "data": rows[1:], "rows": len(rows)})
            if shape.shape_type == 13:  # PICTURE
                picture_count += 1
        notes = ""
        if slide.has_notes_slide:
            notes = slide.notes_slide.notes_text_frame.text.strip()
        title = texts[0] if texts else ""
        slides.append(
            {
                "slide": i,
                "title": title,
                "texts": texts,
                "tables": tables,
                "notes": notes,
                "pictures": picture_count,
                "layout": slide.slide_layout.name,
            }
        )
    cp = prs.core_properties
    return {
        "slides": slides,
        "slide_count": len(slides),
        "metadata": {"title": cp.title or "", "author": cp.author or "", "created": str(cp.created or "")},
        "slide_size": {"width_emu": prs.slide_width, "height_emu": prs.slide_height},
    }


# --------------------------------------------------------------------------
# Excel
# --------------------------------------------------------------------------
def read_workbook(path: Path, *, max_preview_rows: int = 20) -> dict[str, Any]:
    p = Path(path)
    suf = p.suffix.lower()
    if suf == ".xls":
        return _read_workbook_via_xlrd(p, max_preview_rows=max_preview_rows)
    import openpyxl

    wb = openpyxl.load_workbook(str(p), data_only=True)
    try:
        sheets: list[dict[str, Any]] = []
        for ws in wb.worksheets:
            rows = list(ws.iter_rows(values_only=True))
            while rows and all(c is None or str(c).strip() == "" for c in rows[-1]):
                rows.pop()
            header_idx = _guess_header_row(rows)
            header = [("" if c is None else str(c).strip()) for c in rows[header_idx]] if rows else []
            data_rows = rows[header_idx + 1:] if rows else []
            non_empty_rows = sum(1 for r in data_rows if any(c is not None and str(c).strip() != "" for c in r))
            sheets.append(
                {
                    "sheet": ws.title,
                    "dimensions": ws.dimensions,
                    "max_row": ws.max_row,
                    "max_column": ws.max_column,
                    "freeze_panes": ws.freeze_panes,
                    "merged_cells": [str(r) for r in list(ws.merged_cells.ranges)[:50]],
                    "header_row_index": header_idx + 1,
                    "header": header,
                    "data_rows": non_empty_rows,
                    "preview": [
                        [("" if c is None else str(c)) for c in r] for r in data_rows[:max_preview_rows]
                    ],
                    "formulas": _count_formulas(ws, limit=2000),
                    "charts": len(getattr(ws, "_charts", []) or []),
                    "images": len(getattr(ws, "_images", []) or []),
                }
            )
        defined = []
        try:
            defined = [n for n in wb.defined_names]
        except Exception:  # noqa: BLE001
            pass
        return {"sheets": sheets, "sheet_count": len(sheets), "defined_names": defined}
    finally:
        wb.close()


def _count_formulas(ws, limit: int = 2000) -> int:
    n = 0
    for row in ws.iter_rows():
        for cell in row:
            if isinstance(cell.value, str) and cell.value.startswith("="):
                n += 1
                if n >= limit:
                    return n
    return n


def _guess_header_row(rows: list[tuple]) -> int:
    """Excel exports often start with a title row; find the densest text row."""
    best, best_score = 0, -1.0
    for i, row in enumerate(rows[:10]):
        vals = [c for c in row if c is not None and str(c).strip() != ""]
        if not vals:
            continue
        texty = sum(1 for v in vals if not _is_number(v))
        score = len(vals) * 0.6 + texty * 0.6 - i * 0.1
        if score > best_score:
            best, best_score = i, score
    return best


def _is_number(v: Any) -> bool:
    if isinstance(v, (int, float)):
        return True
    try:
        float(str(v).replace(",", "").replace("%", ""))
        return True
    except (TypeError, ValueError):
        return False


def _read_workbook_via_xlrd(path: Path, *, max_preview_rows: int = 20) -> dict[str, Any]:
    import xlrd

    book = xlrd.open_workbook(str(path))
    sheets = []
    for ws in book.sheets():
        rows = [tuple(ws.row_values(r)) for r in range(ws.nrows)]
        header_idx = _guess_header_row(rows)
        header = [str(c).strip() for c in rows[header_idx]] if rows else []
        data_rows = rows[header_idx + 1:]
        sheets.append(
            {
                "sheet": ws.name,
                "dimensions": f"A1:{xlrd.formula.colname(max(ws.ncols - 1, 0))}{ws.nrows}",
                "max_row": ws.nrows,
                "max_column": ws.ncols,
                "header_row_index": header_idx + 1,
                "header": header,
                "data_rows": len(data_rows),
                "preview": [[("" if c is None else str(c)) for c in r] for r in data_rows[:max_preview_rows]],
                "formulas": 0,
                "charts": 0,
                "images": 0,
            }
        )
    return {"sheets": sheets, "sheet_count": len(sheets), "defined_names": []}


# --------------------------------------------------------------------------
# markdown / text rendering of extracted structure
# --------------------------------------------------------------------------
def docx_to_markdown(data: dict[str, Any]) -> str:
    L: list[str] = []
    for b in data["blocks"]:
        t = b["type"]
        if t == "heading":
            L.append("#" * min(int(b.get("level", 1)) + 0, 6) + " " + b["text"].strip())
            L.append("")
        elif t == "list_item":
            L.append(f"- {b['text'].strip()}")
        elif t == "table":
            L.append(_table_md(b["header"], b["data"]))
            L.append("")
        elif t == "blank":
            L.append("")
        else:
            txt = b["text"].strip()
            if txt:
                L.append(txt)
                L.append("")
    return "\n".join(L).strip() + "\n"


def _table_md(header: list[str], rows: list[list[str]]) -> str:
    if not header:
        return ""
    header = [str(h).replace("|", "\\|").replace("\n", " ") for h in header]
    lines = ["| " + " | ".join(header) + " |", "| " + " | ".join("---" for _ in header) + " |"]
    for r in rows:
        cells = [str(c).replace("|", "\\|").replace("\n", " ") for c in r]
        cells = (cells + [""] * len(header))[: len(header)]
        lines.append("| " + " | ".join(cells) + " |")
    return "\n".join(lines)


def pptx_to_markdown(data: dict[str, Any]) -> str:
    L: list[str] = []
    for s in data["slides"]:
        L.append(f"## Slide {s['slide']}" + (f" — {s['title']}" if s["title"] else ""))
        L.append("")
        for txt in s["texts"][1:] if s["title"] else s["texts"]:
            for line in txt.splitlines():
                if line.strip():
                    L.append(f"- {line.strip()}")
            L.append("")
        for t in s["tables"]:
            L.append(_table_md(t["header"], t["data"]))
            L.append("")
        if s["notes"]:
            L.append(f"> 备注: {s['notes']}")
            L.append("")
    return "\n".join(L).strip() + "\n"


def pdf_to_markdown(pages: list[dict[str, Any]], meta: dict[str, Any]) -> str:
    L = [f"<!-- pages: {meta.get('pages')} -->", ""]
    for p in pages:
        L.append(f"## 第 {p['page']} 页")
        L.append("")
        text = p["text"].strip()
        L.append(text if text else "_（本页无可提取文本，可能是扫描图像）_")
        L.append("")
    return "\n".join(L)


# --------------------------------------------------------------------------
# command: extract
# --------------------------------------------------------------------------
DOC_SUFFIXES = {".pdf", ".docx", ".doc", ".pptx", ".ppt", ".xlsx", ".xls", ".xlsm"}


def cmd_extract(args) -> Result:
    res = Result("extract")
    files = resolve_inputs(args.input)
    out = out_dir(getattr(args, "out", None), "extract")
    fmt = (getattr(args, "format", None) or "markdown").lower()
    password = getattr(args, "password", None)
    res.data["documents"] = []

    for f in files:
        suf = f.suffix.lower()
        entry: dict[str, Any] = {"file": str(f.resolve()), "format": suf.lstrip("."), "bytes": f.stat().st_size}
        try:
            if suf == ".pdf":
                pages, meta = read_pdf_text(f, password=password)
                entry["metadata"] = meta
                entry["page_count"] = len(pages)
                entry["word_count"] = pdf_word_count(pages)
                if meta.get("looks_scanned"):
                    res.warn(
                        f"{f.name}: almost no extractable text — it is likely a scan; "
                        f"use `ocr pdf` instead of `extract`"
                    )
                if fmt in ("json", "all"):
                    p = unique_path(out / f"{safe_stem(f.stem)}.pdf.json")
                    write_text(p, json.dumps({"metadata": meta, "pages": pages}, ensure_ascii=False, indent=2))
                    res.add_artifact(p, f"{f.name} page text (json)")
                if fmt in ("markdown", "md", "all"):
                    p = unique_path(out / f"{safe_stem(f.stem)}.md")
                    write_text(p, pdf_to_markdown(pages, meta))
                    res.add_artifact(p, f"{f.name} -> markdown")
                if fmt in ("text", "txt"):
                    p = unique_path(out / f"{safe_stem(f.stem)}.txt")
                    write_text(p, "\n\n".join(f"[page {x['page']}]\n{x['text']}" for x in pages))
                    res.add_artifact(p, f"{f.name} -> text")
                tables = read_pdf_tables(f)
                if tables:
                    entry["tables_found"] = len(tables)
                    tdir = out / "tables"
                    tdir.mkdir(exist_ok=True)
                    for t in tables:
                        tname = unique_path(tdir / f"{safe_stem(f.stem)}_p{t['page']}_t{t['table']}.csv")
                        _write_rows_csv(tname, t["header"], t["data"])
                        res.add_artifact(tname, f"table p{t['page']}#{t['table']}")
                if getattr(args, "images", False):
                    imgs = extract_pdf_images(f, out / "images", password=password)
                    entry["images_extracted"] = len(imgs)
                res.data["documents"].append(entry)

            elif suf in (".docx", ".doc"):
                if suf == ".doc":
                    raise OfficeKitError(
                        f"{f.name} is legacy .doc; convert first with `doc convert --to docx` (uses Word)"
                    )
                data = read_docx(f)
                entry["paragraphs"] = data["paragraphs"]
                entry["table_count"] = len(data["tables"])
                entry["heading_count"] = len(data["headings"])
                entry["characters"] = data["characters"]
                entry["metadata"] = data["metadata"]
                entry["outline"] = [
                    {"level": h.get("level"), "text": h["text"]} for h in data["headings"]
                ]
                if fmt in ("json", "all"):
                    p = unique_path(out / f"{safe_stem(f.stem)}.docx.json")
                    write_text(p, json.dumps(data, ensure_ascii=False, indent=2))
                    res.add_artifact(p, f"{f.name} full structure (json)")
                if fmt in ("markdown", "md", "all"):
                    p = unique_path(out / f"{safe_stem(f.stem)}.md")
                    write_text(p, docx_to_markdown(data))
                    res.add_artifact(p, f"{f.name} -> markdown")
                if fmt in ("text", "txt"):
                    p = unique_path(out / f"{safe_stem(f.stem)}.txt")
                    write_text(p, "\n".join(b.get("text", "") for b in data["blocks"]))
                    res.add_artifact(p, f"{f.name} -> text")
                if getattr(args, "images", False):
                    n = extract_docx_images(f, out / "images")
                    entry["images_extracted"] = n
                res.data["documents"].append(entry)

            elif suf in (".pptx", ".ppt"):
                if suf == ".ppt":
                    raise OfficeKitError(
                        f"{f.name} is legacy .ppt; convert first with `doc convert --to pptx` (uses PowerPoint)"
                    )
                data = read_pptx(f)
                entry["slide_count"] = data["slide_count"]
                entry["metadata"] = data["metadata"]
                entry["outline"] = [{"slide": s["slide"], "title": s["title"]} for s in data["slides"]]
                if fmt in ("json", "all"):
                    p = unique_path(out / f"{safe_stem(f.stem)}.pptx.json")
                    write_text(p, json.dumps(data, ensure_ascii=False, indent=2))
                    res.add_artifact(p, f"{f.name} full structure (json)")
                if fmt in ("markdown", "md", "all"):
                    p = unique_path(out / f"{safe_stem(f.stem)}.md")
                    write_text(p, pptx_to_markdown(data))
                    res.add_artifact(p, f"{f.name} -> markdown")
                if getattr(args, "images", False):
                    n = extract_pptx_images(f, out / "images")
                    entry["images_extracted"] = n
                res.data["documents"].append(entry)

            elif suf in (".xlsx", ".xls", ".xlsm"):
                data = read_workbook(f, max_preview_rows=getattr(args, "preview_rows", 20))
                entry["sheet_count"] = data["sheet_count"]
                entry["sheets"] = [
                    {
                        "sheet": s["sheet"],
                        "rows": s["data_rows"],
                        "columns": len(s["header"]),
                        "header": s["header"],
                        "formulas": s["formulas"],
                        "charts": s["charts"],
                    }
                    for s in data["sheets"]
                ]
                if fmt in ("json", "all"):
                    p = unique_path(out / f"{safe_stem(f.stem)}.xlsx.json")
                    write_text(p, json.dumps(data, ensure_ascii=False, indent=2))
                    res.add_artifact(p, f"{f.name} workbook structure (json)")
                if fmt in ("markdown", "md", "all"):
                    L = [f"# {f.name}", ""]
                    for s in data["sheets"]:
                        L.append(f"## 工作表: {s['sheet']}")
                        L.append("")
                        L.append(f"- 数据行: {s['data_rows']} ・ 列数: {len(s['header'])} ・ 维度: {s['dimensions']}")
                        L.append("")
                        L.append(_table_md(s["header"], s["preview"]))
                        L.append("")
                    p = unique_path(out / f"{safe_stem(f.stem)}.md")
                    write_text(p, "\n".join(L))
                    res.add_artifact(p, f"{f.name} -> markdown")
                if fmt in ("csv", "all"):
                    for s in data["sheets"]:
                        try:
                            df = read_table(f, sheet=s["sheet"], header=max(s["header_row_index"] - 1, 0))
                        except OfficeKitError:
                            continue
                        p = unique_path(out / f"{safe_stem(f.stem)}_{safe_stem(s['sheet'])}.csv")
                        df.to_csv(p, index=False, encoding="utf-8-sig")
                        res.add_artifact(p, f"sheet {s['sheet']} -> csv")
                res.data["documents"].append(entry)

            else:
                entry["error"] = f"unsupported document type: {suf}"
                res.warn(f"{f.name}: unsupported document type, skipped")
                res.data["documents"].append(entry)
        except OfficeKitError as exc:
            entry["error"] = str(exc)
            res.warn(f"{f.name}: {exc}")
            res.data["documents"].append(entry)
    return res


def _write_rows_csv(path: Path, header: list[str], rows: list[list[str]]) -> None:
    import csv

    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", encoding="utf-8-sig", newline="") as fh:
        w = csv.writer(fh)
        w.writerow(header)
        w.writerows(rows)


def extract_docx_images(path: Path, out: Path) -> int:
    """docx is a zip; media lives under word/media."""
    out.mkdir(parents=True, exist_ok=True)
    n = 0
    with zipfile.ZipFile(path) as z:
        for name in z.namelist():
            if name.startswith("word/media/") and not name.endswith("/"):
                data = z.read(name)
                target = unique_path(out / f"{safe_stem(Path(path).stem)}_{Path(name).name}")
                target.write_bytes(data)
                n += 1
    return n


def extract_pptx_images(path: Path, out: Path) -> int:
    out.mkdir(parents=True, exist_ok=True)
    n = 0
    with zipfile.ZipFile(path) as z:
        for name in z.namelist():
            if name.startswith("ppt/media/") and not name.endswith("/"):
                data = z.read(name)
                target = unique_path(out / f"{safe_stem(Path(path).stem)}_{Path(name).name}")
                target.write_bytes(data)
                n += 1
    return n


def extract_xlsx_images(path: Path, out: Path) -> int:
    out.mkdir(parents=True, exist_ok=True)
    n = 0
    with zipfile.ZipFile(path) as z:
        for name in z.namelist():
            if name.startswith("xl/media/") and not name.endswith("/"):
                data = z.read(name)
                target = unique_path(out / f"{safe_stem(Path(path).stem)}_{Path(name).name}")
                target.write_bytes(data)
                n += 1
    return n
