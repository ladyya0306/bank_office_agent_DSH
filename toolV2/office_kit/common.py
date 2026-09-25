"""Shared plumbing for office_kit: encoding, JSON envelope, artifact paths,
table loading, and CJK-safe plotting.

Design rules that every command follows:
  * stdout carries exactly one JSON envelope; nothing else is ever printed.
  * every produced file path is reported in ``artifacts`` so the agent can
    hand it straight to the ``present`` tool.
  * text is written as UTF-8 with BOM where Excel needs to detect it, and
    Chinese is never silently mangled.
"""
from __future__ import annotations

import io
import json
import re
import sys
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any, Iterable, Sequence

# --------------------------------------------------------------------------
# stdout / stderr encoding: Windows consoles default to cp936 and raise
# UnicodeEncodeError on emoji or rare CJK.  Force UTF-8 with replacement.
# --------------------------------------------------------------------------
for _stream_name in ("stdout", "stderr"):
    _s = getattr(sys, _stream_name, None)
    if _s is not None and hasattr(_s, "reconfigure"):
        try:
            _s.reconfigure(encoding="utf-8", errors="replace")
        except Exception:  # noqa: BLE001
            pass


class OfficeKitError(Exception):
    """User-facing failure; the CLI turns this into ok=false with a message."""


# --------------------------------------------------------------------------
# artifact handling
# --------------------------------------------------------------------------
DEFAULT_ARTIFACT_DIR = "_office_out"


@dataclass
class Result:
    """Accumulates the JSON envelope as a command runs."""

    command: str
    data: dict[str, Any] = field(default_factory=dict)
    artifacts: list[dict[str, str]] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)

    def add_artifact(self, path: str | Path, description: str = "") -> str:
        p = Path(path).resolve()
        self.artifacts.append({"path": str(p), "description": description})
        return str(p)

    def warn(self, message: str) -> None:
        self.warnings.append(message)

    def to_json(self, ok: bool = True, error: str | None = None) -> str:
        payload: dict[str, Any] = {
            "ok": ok,
            "command": self.command,
            "data": self.data,
            "artifacts": self.artifacts,
        }
        if self.warnings:
            payload["warnings"] = self.warnings
        if error:
            payload["error"] = error
        return json.dumps(payload, ensure_ascii=False, indent=2, default=_json_default)


def _json_default(o: Any) -> Any:
    if isinstance(o, (datetime,)):
        return o.isoformat(sep=" ")
    if hasattr(o, "item"):  # numpy scalar
        try:
            return o.item()
        except Exception:  # noqa: BLE001
            pass
    if isinstance(o, (set, frozenset)):
        return sorted(o, key=str)
    if isinstance(o, bytes):
        return f"<{len(o)} bytes>"
    if isinstance(o, Path):
        return str(o)
    return str(o)


def out_dir(base: str | Path | None, command: str) -> Path:
    """Resolve the artifact directory and create it.

    **工作区优先**：没显式给 `--out` 时，如果当前目录（或它的上级）是一个**工作区**，
    产物就落到 `<工作区>\\out\\未编号批次\\_报告\\<命令>\\`——
    这样"这次活儿干在哪个文件夹里"和"东西跑哪去了"是同一个答案（[26 §2.9]）。
    找不到工作区才退回原来的 `./_office_out/<command>`（老行为不变）。
    """
    if base:
        root = Path(base).expanduser()
        d = root if root.name == command else root / command
        d.mkdir(parents=True, exist_ok=True)
        return d
    try:
        from .workroot import find_workroot

        found = find_workroot()
        if found:
            d = found[0] / "out" / "未编号批次" / "_报告" / command
            d.mkdir(parents=True, exist_ok=True)
            return d
    except Exception:  # noqa: BLE001 - 找不到工作区就走老路，不许因此让命令失败
        pass
    d = Path.cwd() / DEFAULT_ARTIFACT_DIR / command
    d.mkdir(parents=True, exist_ok=True)
    return d


def safe_stem(name: str, fallback: str = "out") -> str:
    """Make an arbitrary string safe to use as a filename stem."""
    s = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", str(name)).strip(" .")
    s = re.sub(r"\s+", "_", s)
    return (s[:80] or fallback)


# --------------------------------------------------------------------------
# text / encoding helpers
# --------------------------------------------------------------------------
ENCODING_CANDIDATES = ("utf-8-sig", "utf-8", "gb18030", "gbk", "big5", "latin-1")

# Excel and downstream tools detect UTF-8 only when a BOM is present.
CSV_ENCODING = "utf-8-sig"


def read_text(path: str | Path, encoding: str | None = None) -> tuple[str, str]:
    """Read a text file, auto-detecting the encoding. Returns (text, encoding)."""
    p = Path(path)
    if not p.exists():
        raise OfficeKitError(f"file not found: {p}")
    if encoding and encoding.lower() != "auto":
        return p.read_text(encoding=encoding, errors="replace"), encoding
    raw = p.read_bytes()
    for enc in ENCODING_CANDIDATES:
        try:
            return raw.decode(enc), enc
        except UnicodeDecodeError:
            continue
    return raw.decode("utf-8", errors="replace"), "utf-8(replace)"


def write_text(path: str | Path, text: str, encoding: str = "utf-8") -> Path:
    p = Path(path)
    p.parent.mkdir(parents=True, exist_ok=True)
    p.write_text(text, encoding=encoding, newline="")
    return p


def sniff_delimiter(sample: str) -> str:
    counts = {d: sample.count(d) for d in (",", "\t", ";", "|")}
    best = max(counts, key=lambda k: counts[k])
    return best if counts[best] > 0 else ","


# --------------------------------------------------------------------------
# table loading
# --------------------------------------------------------------------------
TABLE_SUFFIXES = {
    ".csv",
    ".tsv",
    ".txt",
    ".xlsx",
    ".xlsm",
    ".xls",
    ".xlsb",
    ".ods",
    ".json",
    ".jsonl",
    ".ndjson",
    ".parquet",
    ".feather",
}


def resolve_inputs(pattern: str | Sequence[str]) -> list[Path]:
    """Accept a file, a directory, or a glob and return sorted existing files.

    Directories are expanded to their *direct* children: this is what table and
    document commands want.  Use :func:`resolve_roots` when the directory itself
    matters (recursive walks).
    """
    import glob as _glob

    items = [pattern] if isinstance(pattern, str) else list(pattern)
    found: list[Path] = []
    for item in items:
        for token in str(item).split(";"):
            token = token.strip()
            if not token:
                continue
            p = Path(token).expanduser()
            if p.is_dir():
                found.extend(sorted(x for x in p.iterdir() if x.is_file()))
            elif p.exists():
                found.append(p)
            else:
                hits = [Path(h) for h in _glob.glob(token, recursive=True)]
                if not hits:
                    raise OfficeKitError(f"no such file or directory: {token}")
                found.extend(sorted(h for h in hits if h.is_file()))
    seen: set[str] = set()
    unique: list[Path] = []
    for p in found:
        key = str(p.resolve()).lower()
        if key not in seen:
            seen.add(key)
            unique.append(p)
    if not unique:
        raise OfficeKitError("no input files matched")
    return unique


def resolve_roots(pattern: str | Sequence[str]) -> list[Path]:
    """Resolve inputs while keeping directories intact (for recursive walks)."""
    import glob as _glob

    items = [pattern] if isinstance(pattern, str) else list(pattern)
    found: list[Path] = []
    for item in items:
        for token in str(item).split(";"):
            token = token.strip()
            if not token:
                continue
            p = Path(token).expanduser()
            if p.exists():
                found.append(p)
                continue
            hits = [Path(h) for h in _glob.glob(token, recursive=True)]
            if not hits:
                raise OfficeKitError(f"no such file or directory: {token}")
            found.extend(hits)
    seen: set[str] = set()
    unique: list[Path] = []
    for p in found:
        key = str(p.resolve()).lower()
        if key not in seen:
            seen.add(key)
            unique.append(p)
    if not unique:
        raise OfficeKitError("no input paths matched")
    return unique


def read_table(
    path: str | Path,
    *,
    sheet: str | int | None = None,
    header: int | None = 0,
    encoding: str | None = None,
    delimiter: str | None = None,
    dtype_backend: bool = True,
    **read_kwargs: Any,
):
    """Read any common tabular file into a pandas DataFrame."""
    import pandas as pd

    p = Path(path)
    if not p.exists():
        raise OfficeKitError(f"file not found: {p}")
    suf = p.suffix.lower()
    kw: dict[str, Any] = {}
    if dtype_backend:
        kw["dtype_backend"] = "numpy_nullable"

    try:
        if suf in (".csv", ".tsv", ".txt"):
            text, enc = read_text(p, encoding)
            if delimiter is None:
                sample = text[:8192]
                delimiter = "\t" if suf == ".tsv" else sniff_delimiter(sample)
            buf = io.StringIO(text)
            return pd.read_csv(
                buf,
                sep=delimiter,
                header=header,
                encoding_errors="replace",
                skip_blank_lines=True,
                **kw,
                **read_kwargs,
            )
        if suf in (".xlsx", ".xlsm", ".xls", ".xlsb", ".ods"):
            engine = None
            if suf == ".xls":
                kw.pop("dtype_backend", None)  # xlrd cannot use nullable backend
                engine = "xlrd"
            elif suf == ".ods":
                kw.pop("dtype_backend", None)
                engine = "odf"
            return pd.read_excel(
                p,
                sheet_name=0 if sheet is None else sheet,
                header=header,
                engine=engine,
                **kw,
                **read_kwargs,
            )
        if suf == ".json":
            return pd.read_json(p, **kw, **read_kwargs)
        if suf in (".jsonl", ".ndjson"):
            return pd.read_json(p, lines=True, **kw, **read_kwargs)
        if suf == ".parquet":
            return pd.read_parquet(p, **read_kwargs)
        if suf == ".feather":
            return pd.read_feather(p, **read_kwargs)
    except OfficeKitError:
        raise
    except Exception as exc:  # noqa: BLE001
        raise OfficeKitError(f"could not read {p.name} as a table: {exc}") from exc
    raise OfficeKitError(
        f"unsupported table format {suf!r}; expected one of {sorted(TABLE_SUFFIXES)}"
    )


def list_sheets(path: str | Path) -> list[str]:
    """Workbook sheet names without loading the data."""
    p = Path(path)
    suf = p.suffix.lower()
    try:
        if suf in (".xlsx", ".xlsm"):
            import openpyxl

            wb = openpyxl.load_workbook(p, read_only=True, data_only=True)
            names = list(wb.sheetnames)
            wb.close()
            return names
        if suf == ".xls":
            import xlrd

            return list(xlrd.open_workbook(p).sheet_names())
        if suf in (".xlsb", ".ods"):
            import pandas as pd

            xls = pd.ExcelFile(p)
            return list(xls.sheet_names)
    except Exception as exc:  # noqa: BLE001
        raise OfficeKitError(f"could not list sheets of {p.name}: {exc}") from exc
    return []


def frame_preview(df, rows: int = 5) -> dict[str, Any]:
    """Compact machine-friendly description of a frame's head."""
    head = df.head(rows)
    return {
        "columns": [str(c) for c in df.columns],
        "rows": [
            {str(k): _cell(v) for k, v in rec.items()}
            for rec in head.to_dict(orient="records")
        ],
    }


def _cell(v: Any) -> Any:
    import pandas as pd

    if v is None or (not isinstance(v, (list, dict, tuple)) and pd.isna(v)):
        return None
    if hasattr(v, "isoformat"):
        return v.isoformat() if not hasattr(v, "hour") else v.isoformat(sep=" ")
    if hasattr(v, "item"):
        try:
            return v.item()
        except Exception:  # noqa: BLE001
            return str(v)
    return v


# --------------------------------------------------------------------------
# plotting with CJK support
# --------------------------------------------------------------------------
CJK_FONT_CANDIDATES = (
    "Microsoft YaHei",
    "SimHei",
    "Noto Sans CJK SC",
    "Source Han Sans SC",
    "Source Han Sans CN",
    "WenQuanYi Zen Hei",
    "PingFang SC",
    "Hiragino Sans GB",
    "DengXian",
    "SimSun",
    "KaiTi",
    "FangSong",
)


def configure_matplotlib():
    """Import pyplot with a CJK-capable font selected, or report what's wrong.

    Without this, Chinese labels render as tofu boxes - a silent, ugly failure.
    """
    import matplotlib

    matplotlib.use("Agg")
    from matplotlib import font_manager, pyplot as plt

    available = {f.name for f in font_manager.fontManager.ttflist}
    chosen = next((c for c in CJK_FONT_CANDIDATES if c in available), None)
    if chosen:
        plt.rcParams["font.sans-serif"] = [chosen, "DejaVu Sans"]
    plt.rcParams["axes.unicode_minus"] = False
    plt.rcParams["figure.autolayout"] = True
    plt.rcParams["savefig.dpi"] = 150
    plt.rcParams["figure.figsize"] = (10, 6)
    return plt, chosen


# --------------------------------------------------------------------------
# color palette shared by charts and reports
# --------------------------------------------------------------------------
PALETTE = [
    "#4C78A8",
    "#F58518",
    "#54A24B",
    "#E45756",
    "#72B7B2",
    "#EECA3B",
    "#B279A2",
    "#FF9DA6",
    "#9D755D",
    "#BAB0AC",
]


def human_bytes(n: float) -> str:
    for unit in ("B", "KB", "MB", "GB", "TB"):
        if abs(n) < 1024 or unit == "TB":
            return f"{n:.1f} {unit}" if unit != "B" else f"{int(n)} B"
        n /= 1024
    return f"{n:.1f} TB"


def ensure_parent(path: str | Path) -> Path:
    p = Path(path)
    p.parent.mkdir(parents=True, exist_ok=True)
    return p


def chunked(seq: Iterable[Any], size: int):
    buf: list[Any] = []
    for item in seq:
        buf.append(item)
        if len(buf) >= size:
            yield buf
            buf = []
    if buf:
        yield buf


def unique_path(path: str | Path) -> Path:
    """Avoid clobbering: append _1, _2, ... before the suffix."""
    p = Path(path)
    if not p.exists():
        return p
    for i in range(1, 10000):
        cand = p.with_name(f"{p.stem}_{i}{p.suffix}")
        if not cand.exists():
            return cand
    raise OfficeKitError(f"cannot find a free filename near {p}")
