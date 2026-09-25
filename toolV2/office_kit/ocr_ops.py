"""OCR: images and scanned PDFs -> text, tables, searchable PDF, Excel.

Engine is RapidOCR (PP-OCR models on ONNXRuntime) which is already installed and
handles Chinese + English out of the box.  Table structure is reconstructed from
the recognised box geometry, so ruled invoice/statement scans survive as grids.
"""
from __future__ import annotations

import csv
import json
import re
from pathlib import Path
from typing import Any, Sequence

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

IMAGE_SUFFIXES = {".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff", ".webp", ".gif", ".jfif", ".pjpeg"}
OCR_SCALE_DEFAULT = 2.5

_ENGINE = None
_CJK_FONT_NAME: str | None = None


def engine():
    """Lazily construct RapidOCR; the first call loads ~15 MB of ONNX models."""
    global _ENGINE
    if _ENGINE is None:
        try:
            from rapidocr_onnxruntime import RapidOCR
        except ImportError as exc:  # pragma: no cover
            raise OfficeKitError(
                f"RapidOCR is not installed ({exc}); run: python -m pip install rapidocr-onnxruntime"
            ) from exc
        try:
            _ENGINE = RapidOCR()
        except Exception as exc:  # noqa: BLE001
            raise OfficeKitError(f"failed to initialise the OCR engine: {exc}") from exc
    return _ENGINE


def engine_info() -> dict[str, Any]:
    info: dict[str, Any] = {"engine": "rapidocr-onnxruntime", "languages": ["ch", "en"]}
    try:
        import rapidocr_onnxruntime  # noqa: F401

        info["installed"] = True
        info["model_loaded"] = _ENGINE is not None
    except ImportError:
        info["installed"] = False
    try:
        import onnxruntime

        info["onnxruntime"] = onnxruntime.__version__
        info["providers"] = onnxruntime.get_available_providers()
    except Exception:  # noqa: BLE001
        pass
    return info


# --------------------------------------------------------------------------
# preprocessing
# --------------------------------------------------------------------------
def preprocess(image, *, mode: str = "auto"):
    """Improve OCR odds on low-quality scans. Returns a PIL image or ndarray."""
    import numpy as np
    from PIL import ImageEnhance, ImageFilter, ImageOps

    if mode == "none":
        return image
    img = image.convert("L") if mode in ("gray", "threshold", "auto") else image.convert("RGB")
    if mode in ("threshold",):
        img = img.point(lambda p: 255 if p > 140 else 0)
    elif mode == "auto":
        arr = np.asarray(img)
        # Heuristic: dark, low-contrast scans benefit from autocontrast + threshold.
        mean = float(arr.mean())
        std = float(arr.std())
        if std < 55 or mean < 110:
            img = ImageOps.autocontrast(img, cutoff=1)
            img = img.filter(ImageFilter.MedianFilter(size=3))
        img = img.convert("RGB")
    elif mode == "enhance":
        img = ImageOps.autocontrast(img.convert("L"), cutoff=1).convert("RGB")
        img = ImageEnhance.Contrast(img).enhance(1.6)
    if img.mode == "L":
        img = img.convert("RGB")
    return img


def upscale_for_ocr(image, *, min_height: int = 1200, max_scale: float = 3.0):
    from PIL import Image

    w, h = image.size
    if h >= min_height:
        return image
    scale = min(min_height / max(h, 1), max_scale)
    if scale <= 1.05:
        return image
    return image.resize((int(w * scale), int(h * scale)), Image.LANCZOS)


# --------------------------------------------------------------------------
# core recognition
# --------------------------------------------------------------------------
def ocr_image(path: Path, *, pre: str = "auto", min_height: int = 1200) -> dict[str, Any]:
    """Recognise one image. Returns lines with text, score and pixel box."""
    from PIL import Image

    if not path.exists():
        raise OfficeKitError(f"image not found: {path}")
    try:
        image = Image.open(path)
        image.load()
    except Exception as exc:  # noqa: BLE001
        raise OfficeKitError(f"cannot open image {path.name}: {exc}") from exc

    original_size = list(image.size)
    image = upscale_for_ocr(image, min_height=min_height)
    prepared = preprocess(image, mode=pre)
    import numpy as np

    arr = np.asarray(prepared)
    result, elapse = engine()(arr)
    lines: list[dict[str, Any]] = []
    for item in result or []:
        box, text, score = item[0], item[1], item[2]
        xs = [float(p[0]) for p in box]
        ys = [float(p[1]) for p in box]
        lines.append(
            {
                "text": str(text),
                "score": round(float(score), 4),
                "box": [[round(float(p[0]), 1), round(float(p[1]), 1)] for p in box],
                "x0": min(xs), "x1": max(xs), "y0": min(ys), "y1": max(ys),
                "cx": sum(xs) / len(xs), "cy": sum(ys) / len(ys),
            }
        )
    lines.sort(key=lambda ln: (round(ln["cy"] / 12), ln["x0"]))
    confidences = [ln["score"] for ln in lines]
    return {
        "file": str(path.resolve()),
        "engine": "rapidocr-onnxruntime",
        "original_size": original_size,
        "ocr_size": list(prepared.size),
        "preprocess": pre,
        "line_count": len(lines),
        "mean_confidence": round(sum(confidences) / len(confidences), 4) if confidences else None,
        "low_confidence_lines": sum(1 for c in confidences if c < 0.6),
        "lines": lines,
        "text": "\n".join(ln["text"] for ln in lines),
        "elapse": elapse,
    }


def group_lines_into_rows(lines: list[dict[str, Any]], *, tol_ratio: float = 0.6) -> list[list[dict[str, Any]]]:
    """Cluster OCR lines into visual rows using vertical overlap of boxes."""
    if not lines:
        return []
    heights = [max(ln["y1"] - ln["y0"], 1.0) for ln in lines]
    median_h = sorted(heights)[len(heights) // 2]
    tol = median_h * tol_ratio
    ordered = sorted(lines, key=lambda ln: (ln["cy"], ln["x0"]))
    rows: list[list[dict[str, Any]]] = []
    for line in ordered:
        placed = False
        for row in rows:
            row_cy = sum(x["cy"] for x in row) / len(row)
            if abs(line["cy"] - row_cy) <= tol:
                row.append(line)
                placed = True
                break
        if not placed:
            rows.append([line])
    for row in rows:
        row.sort(key=lambda ln: ln["x0"])
    rows.sort(key=lambda r: min(ln["cy"] for ln in r))
    return rows


def reconstruct_table(lines: list[dict[str, Any]], *, min_cols: int = 2, min_rows: int = 2) -> dict[str, Any] | None:
    """Turn OCR boxes into a grid by clustering column x-positions.

    Works well for ruled forms, invoices and statements; returns None when the
    layout clearly is not tabular, so callers can fall back to plain text.
    """
    rows = group_lines_into_rows(lines)
    rows = [r for r in rows if len(r) >= 1]
    if len(rows) < min_rows:
        return None
    max_cols = max(len(r) for r in rows)
    if max_cols < min_cols:
        return None

    # Cluster all x-centres to infer the column grid.
    centres = sorted(ln["cx"] for r in rows for ln in r)
    if not centres:
        return None
    widths = [max(ln["x1"] - ln["x0"], 1.0) for r in rows for ln in r]
    median_w = sorted(widths)[len(widths) // 2]
    gap = median_w * 0.8
    clusters: list[list[float]] = [[centres[0]]]
    for c in centres[1:]:
        if c - clusters[-1][-1] > gap:
            clusters.append([c])
        else:
            clusters[-1].append(c)
    anchors = [sum(cl) / len(cl) for cl in clusters]
    if len(anchors) < min_cols:
        return None

    grid: list[list[str]] = []
    for row in rows:
        cells = [""] * len(anchors)
        for line in row:
            idx = min(range(len(anchors)), key=lambda i: abs(anchors[i] - line["cx"]))
            cells[idx] = (cells[idx] + " " + line["text"]).strip()
        grid.append(cells)
    filled = sum(1 for r in grid for c in r if c)
    density = filled / (len(grid) * len(anchors))
    if density < 0.25:
        return None
    return {
        "columns": len(anchors),
        "rows": len(grid),
        "column_anchors": [round(a, 1) for a in anchors],
        "density": round(density, 3),
        "header": grid[0],
        "data": grid[1:],
        "grid": grid,
    }


def lines_to_text(lines: list[dict[str, Any]], *, layout: str = "auto") -> str:
    if layout == "rows":
        rows = group_lines_into_rows(lines)
        return "\n".join("  ".join(ln["text"] for ln in r) for r in rows)
    return "\n".join(ln["text"] for ln in lines)


# --------------------------------------------------------------------------
# PDF OCR
# --------------------------------------------------------------------------
def ocr_pdf_pages(path: Path, *, pages: Sequence[int] | None = None, scale: float = OCR_SCALE_DEFAULT,
                  pre: str = "auto", dpi_note: bool = True) -> list[dict[str, Any]]:
    """Rasterise then OCR a PDF. Returns one entry per page with text and lines."""
    import tempfile

    from .doc_read import render_pdf_pages

    with tempfile.TemporaryDirectory(prefix="officekit_ocr_") as td:
        tmp = Path(td)
        rendered = render_pdf_pages(path, tmp, pages=list(pages) if pages else None, scale=scale, fmt="png")
        if not rendered:
            raise OfficeKitError(f"no pages rendered from {path.name}")
        out: list[dict[str, Any]] = []
        for item in rendered:
            res = ocr_image(Path(item["file"]), pre=pre)
            out.append(
                {
                    "page": item["page"],
                    "text": res["text"],
                    "lines": res["lines"],
                    "line_count": res["line_count"],
                    "mean_confidence": res["mean_confidence"],
                    "render_scale": scale,
                }
            )
    return out


# --------------------------------------------------------------------------
# command: ocr
# --------------------------------------------------------------------------
def cmd_ocr(args) -> Result:
    res = Result("ocr")
    files = resolve_inputs(args.input)
    out = out_dir(getattr(args, "out", None), "ocr")
    pre = getattr(args, "preprocess", "auto") or "auto"
    layout = getattr(args, "layout", "auto") or "auto"
    res.data["engine"] = engine_info()
    res.data["items"] = []

    for f in files:
        suf = f.suffix.lower()
        try:
            if suf == ".pdf":
                pages = ocr_pdf_pages(
                    f,
                    pages=_parse_pages(getattr(args, "pages", None)),
                    scale=float(getattr(args, "scale", OCR_SCALE_DEFAULT) or OCR_SCALE_DEFAULT),
                    pre=pre,
                )
                entry = {
                    "file": str(f.resolve()),
                    "kind": "pdf",
                    "page_count": len(pages),
                    "total_lines": sum(p["line_count"] for p in pages),
                    "mean_confidence": _mean([p["mean_confidence"] for p in pages if p["mean_confidence"]]),
                }
                res.data["items"].append(entry)
                stem = safe_stem(f.stem)
                if getattr(args, "emit_text", True):
                    text = "\n\n".join(f"===== 第 {p['page']} 页 =====\n{p['text']}" for p in pages)
                    p_txt = unique_path(out / f"{stem}_ocr.txt")
                    write_text(p_txt, text)
                    res.add_artifact(p_txt, "OCR text")
                if getattr(args, "emit_json", True):
                    p_json = unique_path(out / f"{stem}_ocr.json")
                    write_text(p_json, json.dumps(pages, ensure_ascii=False, indent=2, default=str))
                    res.add_artifact(p_json, "OCR detail (lines + boxes)")
                if getattr(args, "emit_word", False):
                    from .doc_build import build_docx_from_blocks

                    blocks: list[dict[str, Any]] = []
                    for p in pages:
                        blocks.append({"type": "heading", "level": 2, "text": f"第 {p['page']} 页"})
                        for para in _split_paras(p["text"]):
                            blocks.append({"type": "paragraph", "text": para})
                    p_docx = unique_path(out / f"{stem}_ocr.docx")
                    build_docx_from_blocks(blocks, p_docx, title=f.stem)
                    res.add_artifact(p_docx, "OCR result as Word")
                if getattr(args, "emit_searchable_pdf", False):
                    _try_searchable_pdf(f, pages, unique_path(out / f"{stem}_searchable.pdf"), res)
                if getattr(args, "tables", True) and pages:
                    _emit_tables(res, out, stem, pages)
            elif suf in IMAGE_SUFFIXES:
                data = ocr_image(f, pre=pre)
                entry = {
                    "file": str(f.resolve()),
                    "kind": "image",
                    "line_count": data["line_count"],
                    "mean_confidence": data["mean_confidence"],
                    "low_confidence_lines": data["low_confidence_lines"],
                    "original_size": data["original_size"],
                }
                table = None
                if getattr(args, "tables", True):
                    table = reconstruct_table(data["lines"])
                    entry["table_detected"] = table is not None
                    if table:
                        entry["table_shape"] = [table["rows"], table["columns"]]
                res.data["items"].append(entry)
                stem = safe_stem(f.stem)
                text = lines_to_text(data["lines"], layout=layout)
                if getattr(args, "emit_text", True):
                    p_txt = unique_path(out / f"{stem}_ocr.txt")
                    write_text(p_txt, text)
                    res.add_artifact(p_txt, "OCR text")
                if getattr(args, "emit_markdown", False):
                    p_md = unique_path(out / f"{stem}_ocr.md")
                    md = f"# {f.name} OCR\n\n"
                    md += _table_md(table) if table else text
                    write_text(p_md, md)
                    res.add_artifact(p_md, "OCR markdown")
                if getattr(args, "emit_json", True):
                    p_json = unique_path(out / f"{stem}_ocr.json")
                    write_text(p_json, json.dumps({k: v for k, v in data.items() if k != "elapse"},
                                                  ensure_ascii=False, indent=2, default=str))
                    res.add_artifact(p_json, "OCR detail (lines + boxes)")
                if table:
                    p_csv = unique_path(out / f"{stem}_table.csv")
                    _write_grid_csv(p_csv, table["grid"])
                    res.add_artifact(p_csv, "reconstructed table (csv)")
                    if getattr(args, "emit_excel", True):
                        p_xlsx = unique_path(out / f"{stem}_table.xlsx")
                        _write_grid_xlsx(p_xlsx, table["grid"])
                        res.add_artifact(p_xlsx, "reconstructed table (xlsx)")
                elif getattr(args, "tables", True) and data["line_count"] > 3:
                    res.warn(
                        f"{f.name}: layout did not look tabular, so no grid was built; "
                        f"use `--layout rows` text output instead"
                    )
                if getattr(args, "emit_word", False):
                    from .doc_build import build_docx_from_blocks

                    blocks = [
                        {"type": "paragraph", "text": para} for para in _split_paras(text)
                    ] or [{"type": "paragraph", "text": text}]
                    p_docx = unique_path(out / f"{stem}_ocr.docx")
                    build_docx_from_blocks(blocks, p_docx, title=f.stem)
                    res.add_artifact(p_docx, "OCR result as Word")
                if getattr(args, "emit_searchable_pdf", False):
                    _try_searchable_pdf(
                        f,
                        [{"lines": data["lines"], "text": text, "page": 1}],
                        unique_path(out / f"{stem}_searchable.pdf"),
                        res,
                    )
            else:
                res.warn(f"unsupported input for OCR: {f.name}")
                continue
            if entry.get("mean_confidence") is not None and entry["mean_confidence"] < 0.75:
                res.warn(
                    f"{f.name}: mean confidence {entry['mean_confidence']:.2f} is low — "
                    f"check the recognised text before relying on it"
                )
        except OfficeKitError as exc:
            res.warn(f"{f.name}: {exc}")
            res.data["items"].append({"file": str(f.resolve()), "error": str(exc)})

    if not res.data["items"]:
        raise OfficeKitError("nothing could be processed; check the input files")
    return res


def _parse_pages(spec: str | None) -> list[int] | None:
    if not spec:
        return None
    out: list[int] = []
    for part in str(spec).split(","):
        part = part.strip()
        m = re.match(r"^(\d+)\s*-\s*(\d+)$", part)
        if m:
            out.extend(range(int(m.group(1)), int(m.group(2)) + 1))
        elif part.isdigit():
            out.append(int(part))
    return out or None


def _mean(vals: list[float]) -> float | None:
    return round(sum(vals) / len(vals), 4) if vals else None


def _split_paras(text: str) -> list[str]:
    return [p.strip() for p in re.split(r"\n\s*\n", text) if p.strip()] or []


def _emit_tables(res: Result, out: Path, stem: str, pages: list[dict[str, Any]]) -> None:
    for p in pages:
        table = reconstruct_table(p["lines"])
        if not table:
            continue
        p_csv = unique_path(out / f"{stem}_p{p['page']}_table.csv")
        _write_grid_csv(p_csv, table["grid"])
        res.add_artifact(p_csv, f"page {p['page']} reconstructed table (csv)")


def _write_grid_csv(path: Path, grid: list[list[str]]) -> None:
    ensure_parent(path)
    with open(path, "w", encoding="utf-8-sig", newline="") as fh:
        csv.writer(fh).writerows(grid)


def _write_grid_xlsx(path: Path, grid: list[list[str]]) -> None:
    import xlsxwriter

    ensure_parent(path)
    wb = xlsxwriter.Workbook(str(path))
    ws = wb.add_worksheet("ocr")
    fmt = wb.add_format({"bold": True, "bg_color": "#1F3864", "font_color": "white", "border": 1})
    for r, row in enumerate(grid):
        for c, val in enumerate(row):
            if r == 0:
                ws.write(r, c, val, fmt)
            else:
                ws.write(r, c, val)
    ncols = max((len(r) for r in grid), default=1)
    for c in range(ncols):
        width = max([len(str(grid[r][c])) * (2 if any(ord(ch) > 0x2E80 for ch in str(grid[r][c])) else 1)
                     for r in range(len(grid)) if c < len(grid[r])] or [10])
        ws.set_column(c, c, min(max(width + 2, 8), 50))
    wb.close()


def _table_md(table: dict[str, Any]) -> str:
    grid = table["grid"]
    if not grid:
        return ""
    ncols = len(grid[0])
    lines = ["| " + " | ".join(str(c) for c in grid[0]) + " |",
             "| " + " | ".join("---" for _ in range(ncols)) + " |"]
    for row in grid[1:]:
        cells = (list(row) + [""] * ncols)[:ncols]
        lines.append("| " + " | ".join(str(c).replace("|", "\\|") for c in cells) + " |")
    return "\n".join(lines)


def _register_cjk_font() -> str | None:
    """Register a TTF that actually contains CJK glyphs with reportlab.

    reportlab's built-in Type-1 fonts (Helvetica et al.) have no CJK coverage, so
    an invisible layer drawn with them extracts as tofu boxes.  Registered once
    per process and cached.
    """
    global _CJK_FONT_NAME
    if _CJK_FONT_NAME is not None:
        return _CJK_FONT_NAME
    candidates = [
        r"C:\Windows\Fonts\msyh.ttc",
        r"C:\Windows\Fonts\msyh.ttf",
        r"C:\Windows\Fonts\simhei.ttf",
        r"C:\Windows\Fonts\simsun.ttc",
        r"C:\Windows\Fonts\Deng.ttf",
        "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
        "/System/Library/Fonts/PingFang.ttc",
    ]
    try:
        from reportlab.pdfbase import pdfmetrics
        from reportlab.pdfbase.ttfonts import TTFont
    except ImportError:
        return None
    for path in candidates:
        if not Path(path).exists():
            continue
        try:
            # .ttc collections expose subfont index 0 as the regular face.
            pdfmetrics.registerFont(TTFont("OfficeKitCJK", path, subfontIndex=0))
            _CJK_FONT_NAME = "OfficeKitCJK"
            return _CJK_FONT_NAME
        except Exception:  # noqa: BLE001
            continue
    return None


def _try_searchable_pdf(src: Path, pages: list[dict[str, Any]], target: Path, res: Result) -> None:
    """Attach an invisible text layer so an image or a PDF becomes searchable.

    The recognised box coordinates are reused directly, so the invisible text
    lands exactly on the pixels it came from.  Only the first page of a
    multi-page input is searchable per call; callers pass one page at a time.
    """
    try:
        from reportlab.pdfgen import canvas  # type: ignore
    except ImportError as exc:
        res.warn(f"searchable PDF needs reportlab ({exc}); skipped")
        return
    try:
        import tempfile

        from PIL import Image

        page_data = pages[0] if pages else {"lines": []}
        is_pdf = src.suffix.lower() == ".pdf"
        font_name = _register_cjk_font()
        if font_name is None:
            res.warn(
                "no CJK TrueType font could be registered with reportlab; "
                "the searchable layer cannot render Chinese (skipped)"
            )
            return
        with tempfile.TemporaryDirectory(prefix="officekit_ocr_pdf_") as td:
            tmp = Path(td)
            if is_pdf:
                from .doc_read import render_pdf_pages

                rendered = render_pdf_pages(src, tmp, scale=OCR_SCALE_DEFAULT, fmt="png")
                if not rendered:
                    res.warn("searchable PDF skipped: nothing rendered from the input")
                    return
                raster = Path(rendered[0]["file"])
            else:
                raster = src

            with Image.open(raster) as im:
                w_pt, h_pt = im.size[0] * 0.75, im.size[1] * 0.75
                rgb = im.convert("RGB")
                img_copy = tmp / "raster.png"
                rgb.save(img_copy)

            # Build one PDF: the page image, then the invisible text on top. This
            # avoids pypdf's merge_page ending up with a truncated content stream
            # on the blank-page source produced by pypdfium2.
            ensure_parent(target)
            c = canvas.Canvas(str(target), pagesize=(w_pt, h_pt))
            c.drawImage(str(img_copy), 0, 0, width=w_pt, height=h_pt)
            for line in page_data.get("lines", []):
                size = max((line["y1"] - line["y0"]) * 0.72, 5)
                try:
                    c.setFont(font_name, size)
                    c.setFillAlpha(0)
                except Exception:  # noqa: BLE001
                    pass
                c.drawString(line["x0"] * 0.75, h_pt - line["y1"] * 0.75, line["text"])
            c.save()
        res.add_artifact(target, "searchable PDF (invisible text layer)")
    except Exception as exc:  # noqa: BLE001
        res.warn(f"searchable PDF generation failed: {exc}")
