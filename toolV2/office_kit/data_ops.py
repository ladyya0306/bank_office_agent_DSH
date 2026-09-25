"""Data analysis & reporting: inspect, clean, profile, pivot, chart, report, merge, compare.

Everything here works on pandas DataFrames and writes UTF-8-SIG CSV so that
Excel on a Chinese Windows opens them without mojibake.
"""
from __future__ import annotations

import json
import math
import re
from pathlib import Path
from typing import Any

import numpy as np
import pandas as pd

from .common import (
    CSV_ENCODING,
    PALETTE,
    OfficeKitError,
    Result,
    configure_matplotlib,
    frame_preview,
    list_sheets,
    out_dir,
    read_table,
    resolve_inputs,
    safe_stem,
    unique_path,
    write_text,
)

# --------------------------------------------------------------------------
# shared vocabulary
# --------------------------------------------------------------------------
NULL_TOKENS = {
    "", "na", "n/a", "nan", "null", "none", "nil", "-", "--", "---", "?",
    "#n/a", "#null!", "#div/0!", "无", "空", "未知", "缺失", "\\n", "n.a.",
}

# Chinese full-width digits/punctuation that break numeric parsing.
FULLWIDTH_MAP = {ord(c): ord(c) - 0xFEE0 for c in "０１２３４５６７８９"}
FULLWIDTH_MAP.update({
    ord("．"): ord("."),
    ord("－"): ord("-"),
    ord("＋"): ord("+"),
    ord("，"): ord(","),
    ord("％"): ord("%"),
    ord("　"): ord(" "),
    ord("（"): ord("("),
    ord("）"): ord(")"),
    ord("："): ord(":"),
})

CURRENCY_RE = re.compile(r"[¥$€£￥\s]|RMB|CNY|USD|EUR|GBP", re.IGNORECASE)
PERCENT_RE = re.compile(r"%$")
THOUSAND_RE = re.compile(r"(?<=\d),(?=\d{3}\b)")
BRACKET_NEG_RE = re.compile(r"^\((.*)\)$")


def _is_null_token(v: Any) -> bool:
    if v is None:
        return True
    if isinstance(v, float) and math.isnan(v):
        return True
    try:
        if pd.isna(v):
            return True
    except (TypeError, ValueError):
        pass
    return str(v).strip().lower() in NULL_TOKENS


def _normalize_name(name: str) -> str:
    """Readable, stable, Excel-friendly column name."""
    s = str(name).strip().replace("\n", " ").replace("\r", " ")
    s = re.sub(r"\s+", " ", s)
    s = {"（": "(", "）": ")", "　": " "}.get(s, s) if len(s) == 1 else s
    s = s.replace("（", "(").replace("）", ")").replace("　", " ")
    s = re.sub(r"[^\w\u4e00-\u9fff ()./%\-]+", "_", s)
    s = re.sub(r"\s+", "_", s.strip())
    s = re.sub(r"_+", "_", s).strip("_")
    return s or "column"


def _dedupe_names(names: list[str]) -> list[str]:
    seen: dict[str, int] = {}
    out: list[str] = []
    for n in names:
        if n in seen:
            seen[n] += 1
            out.append(f"{n}_{seen[n]}")
        else:
            seen[n] = 0
            out.append(n)
    return out


def _to_number(series: pd.Series) -> pd.Series:
    """Parse messy numeric text: currency, thousands separators, percent, (neg)."""
    s = series.astype("string")
    s = s.str.translate(FULLWIDTH_MAP)
    s = s.str.strip()
    negative_bracket = s.str.match(BRACKET_NEG_RE, na=False)
    s = s.str.replace(BRACKET_NEG_RE, r"\1", regex=True)
    pct = s.str.contains(PERCENT_RE, na=False)
    s = s.str.replace(PERCENT_RE, "", regex=True)
    s = s.str.replace(CURRENCY_RE, "", regex=True)
    s = s.str.replace(THOUSAND_RE, "", regex=True)
    s = s.str.replace(r"^\((.*)\)$", r"-\1", regex=True)
    num = pd.to_numeric(s, errors="coerce")
    num = num.where(~pct, num)  # percentages stay in percent units
    num = num.mask(negative_bracket & num.notna(), -num.abs())
    return num


def _looks_numeric(series: pd.Series, threshold: float = 0.8) -> bool:
    s = series.dropna()
    if s.empty:
        return False
    if pd.api.types.is_numeric_dtype(s):
        return True
    parsed = _to_number(s.astype("string"))
    return float(parsed.notna().mean()) >= threshold


def _to_datetime(series: pd.Series) -> pd.Series:
    s = series.astype("string").str.strip().str.translate(FULLWIDTH_MAP)
    s = s.str.replace(r"^(\d{4})[年./](\d{1,2})[月./](\d{1,2})日?$", r"\1-\2-\3", regex=True)
    s = s.str.replace(r"^(\d{4})[年./](\d{1,2})[月]?$", r"\1-\2-01", regex=True)
    s = s.str.replace(r"(\d{4})年(\d{1,2})月(\d{1,2})日", r"\1-\2-\3", regex=True)
    out = pd.to_datetime(s, errors="coerce", format="mixed")
    return out


def _looks_datetime(series: pd.Series, threshold: float = 0.8) -> bool:
    if pd.api.types.is_datetime64_any_dtype(series):
        return True
    s = series.dropna()
    if s.empty or pd.api.types.is_numeric_dtype(s):
        return False
    sample = s.head(2000)
    # Date-ish strings only: require a separator or CJK date marker.
    if not sample.astype("string").str.contains(r"[-/年月.]", regex=True, na=False).any():
        return False
    parsed = _to_datetime(sample)
    return float(parsed.notna().mean()) >= threshold


# --------------------------------------------------------------------------
# inspect
# --------------------------------------------------------------------------
def cmd_inspect(args) -> Result:
    res = Result("inspect")
    files = resolve_inputs(args.input)
    res.data["files"] = []
    for f in files:
        entry: dict[str, Any] = {
            "path": str(f.resolve()),
            "bytes": f.stat().st_size,
            "format": f.suffix.lower().lstrip("."),
        }
        try:
            if f.suffix.lower() in (".xlsx", ".xlsm", ".xls", ".xlsb", ".ods"):
                entry["sheets"] = list_sheets(f)
        except OfficeKitError:
            entry["sheets"] = []
        if f.suffix.lower() in (".csv", ".tsv", ".txt"):
            head, enc = _peek_text(f)
            entry["encoding"] = enc
            entry["first_lines"] = head
        try:
            df = read_table(
                f,
                sheet=getattr(args, "sheet", None),
                header=getattr(args, "header", 0),
            )
            entry["table"] = _table_shape(df)
            entry["preview"] = frame_preview(df, rows=getattr(args, "preview_rows", 5))
            entry["problem_columns"] = _quick_problems(df)
        except OfficeKitError as exc:
            entry["table_error"] = str(exc)
        res.data["files"].append(entry)
    return res


def _peek_text(p: Path, lines: int = 5) -> tuple[list[str], str]:
    from .common import read_text

    text, enc = read_text(p)
    return text.splitlines()[:lines], enc


def _table_shape(df: pd.DataFrame) -> dict[str, Any]:
    return {
        "rows": int(len(df)),
        "columns": int(df.shape[1]),
        "column_names": [str(c) for c in df.columns],
        "dtypes": {str(c): str(t) for c, t in df.dtypes.items()},
        "memory_bytes": int(df.memory_usage(deep=True).sum()),
    }


def _quick_problems(df: pd.DataFrame) -> list[dict[str, Any]]:
    problems: list[dict[str, Any]] = []
    for col in df.columns:
        s = df[col]
        nulls = int(s.isna().sum())
        blanks = int((s.astype("string").str.strip() == "").sum())
        if nulls or blanks:
            problems.append(
                {
                    "column": str(col),
                    "issue": "missing",
                    "nulls": nulls,
                    "blank_strings": blanks,
                    "rate": round((nulls + blanks) / max(len(df), 1), 4),
                }
            )
        dup_col = int(s.duplicated().sum())
        if len(df) and dup_col / len(df) > 0.9 and s.nunique(dropna=False) > 1:
            problems.append(
                {"column": str(col), "issue": "mostly_duplicate", "unique": int(s.nunique())}
            )
        if pd.api.types.is_object_dtype(s) or pd.api.types.is_string_dtype(s):
            txt = s.dropna().astype("string")
            if not txt.empty:
                if txt.str.startswith(" ").any() or txt.str.endswith(" ").any():
                    problems.append({"column": str(col), "issue": "untrimmed_whitespace"})
                if txt.str.contains(r"[\uFF01-\uFF5E]", regex=True, na=False).any():
                    problems.append({"column": str(col), "issue": "fullwidth_characters"})
        if s.duplicated().sum() == 0 and len(df) > 1:
            problems.append({"column": str(col), "issue": "all_values_unique"})
    return problems


# --------------------------------------------------------------------------
# clean
# --------------------------------------------------------------------------
def cmd_clean(args) -> Result:
    res = Result("clean")
    src = resolve_inputs(args.input)[0]
    df = read_table(src, sheet=getattr(args, "sheet", None), header=getattr(args, "header", 0))
    original_shape = df.shape
    log: list[dict[str, Any]] = []

    # 1. column names
    if not getattr(args, "keep_column_names", False):
        new_names = _dedupe_names([_normalize_name(c) for c in df.columns])
        renamed = {
            str(o): n for o, n in zip(df.columns, new_names) if str(o) != n
        }
        df.columns = new_names
        if renamed:
            log.append({"step": "rename_columns", "changed": renamed})

    # 2. drop fully empty rows/cols
    if getattr(args, "drop_empty", True):
        before = df.shape
        df = df.dropna(axis=0, how="all").dropna(axis=1, how="all")
        if df.shape != before:
            log.append(
                {
                    "step": "drop_empty",
                    "rows_removed": int(before[0] - df.shape[0]),
                    "cols_removed": int(before[1] - df.shape[1]),
                }
            )

    # 3. drop unnamed/index-like columns created by Excel exports
    unnamed = [c for c in df.columns if re.fullmatch(r"(unnamed[:_ ]?\d*|index|序号)", str(c), re.I)]
    if unnamed and getattr(args, "drop_unnamed", True):
        df = df.drop(columns=unnamed)
        log.append({"step": "drop_unnamed_columns", "columns": [str(c) for c in unnamed]})

    # 4. trim text and normalise null tokens
    text_cols = [
        c for c in df.columns
        if not pd.api.types.is_numeric_dtype(df[c])
        and not pd.api.types.is_datetime64_any_dtype(df[c])
    ]
    trimmed = 0
    nulled = 0
    for c in text_cols:
        s = df[c].astype("string")
        stripped = s.str.replace(r"^\s+|\s+$", "", regex=True)
        stripped = stripped.str.replace(r"\s+", " ", regex=True)
        if not s.equals(stripped):
            trimmed += 1
        mask = stripped.str.lower().isin({t for t in NULL_TOKENS if t})
        if mask.any():
            nulled += int(mask.sum())
            stripped = stripped.mask(mask)
        df[c] = stripped
    if trimmed or nulled:
        log.append({"step": "trim_and_null_tokens", "columns": trimmed, "values_nulled": nulled})

    # 5. full-width -> half-width for text columns
    if getattr(args, "normalize_fullwidth", True):
        fw = 0
        for c in text_cols:
            if c not in df.columns:
                continue
            s = df[c].astype("string")
            conv = s.str.translate(FULLWIDTH_MAP)
            if not s.equals(conv):
                fw += 1
                df[c] = conv
        if fw:
            log.append({"step": "normalize_fullwidth", "columns": fw})

    # 6. type inference
    type_changes: dict[str, str] = {}
    for c in list(df.columns):
        s = df[c]
        if pd.api.types.is_numeric_dtype(s) or pd.api.types.is_datetime64_any_dtype(s):
            continue
        if _looks_datetime(s):
            df[c] = _to_datetime(s)
            type_changes[str(c)] = "datetime"
        elif _looks_numeric(s):
            df[c] = _to_number(s)
            type_changes[str(c)] = "number"
        else:
            df[c] = s.astype("string")
            type_changes[str(c)] = "text"
    if type_changes:
        log.append({"step": "infer_types", "types": type_changes})

    # 7. unify near-duplicate category spellings
    if getattr(args, "fuzzy_categories", True):
        cat_fixes: dict[str, dict[str, str]] = {}
        for c in df.columns:
            if pd.api.types.is_numeric_dtype(df[c]) or pd.api.types.is_datetime64_any_dtype(df[c]):
                continue
            uniq = df[c].dropna().unique()
            if not 2 <= len(uniq) <= 200:
                continue
            canon: dict[str, str] = {}
            for v in uniq:
                key = re.sub(r"[\s_\-()（）\[\]]+", "", str(v)).lower()
                if key in canon and canon[key] != v:
                    cat_fixes.setdefault(str(c), {})[str(v)] = str(canon[key])
                else:
                    canon[key] = str(v)
            if str(c) in cat_fixes:
                df[c] = df[c].astype("string").replace(cat_fixes[str(c)])
        if cat_fixes:
            log.append({"step": "unify_categories", "mapping": cat_fixes})

    # 8. outliers
    if getattr(args, "outliers", "flag") != "none":
        flags: list[str] = []
        for c in df.select_dtypes(include=[np.number]).columns:
            s = df[c]
            if s.dropna().nunique() < 5:
                continue
            q1, q3 = s.quantile(0.25), s.quantile(0.75)
            iqr = q3 - q1
            if iqr <= 0:
                continue
            lo, hi = q1 - 1.5 * iqr, q3 + 1.5 * iqr
            mask = (s < lo) | (s > hi)
            if not mask.any():
                continue
            flag_col = f"{c}_is_outlier"
            df[flag_col] = mask.fillna(False).astype(bool)
            flags.append(flag_col)
            if getattr(args, "outliers", "flag") == "clip":
                df[c] = s.clip(lo, hi)
        if flags:
            log.append({"step": "outliers", "mode": getattr(args, "outliers", "flag"), "columns": flags})

    # 9. dedupe
    if getattr(args, "dedupe", True):
        subset = getattr(args, "dedupe_keys", None)
        if subset:
            keys = [k.strip() for k in str(subset).split(",") if k.strip()]
            missing = [k for k in keys if k not in df.columns]
            if missing:
                res.warn(f"dedupe keys not found, ignoring: {missing}")
                keys = [k for k in keys if k in df.columns] or None
            subset = keys or None
        before = len(df)
        df = df.drop_duplicates(subset=subset, keep="first")
        if before != len(df):
            log.append({"step": "deduplicate", "rows_removed": int(before - len(df)), "keys": subset})

    # 10. column-name override from rules file
    if getattr(args, "rules", None):
        rules = json.loads(Path(args.rules).read_text(encoding="utf-8"))
        rename_map = rules.get("rename_columns", {})
        if rename_map:
            df = df.rename(columns={k: v for k, v in rename_map.items() if k in df.columns})
            log.append({"step": "apply_rules_rename", "mapping": rename_map})
        drop_cols = [c for c in rules.get("drop_columns", []) if c in df.columns]
        if drop_cols:
            df = df.drop(columns=drop_cols)
            log.append({"step": "apply_rules_drop", "columns": drop_cols})

    out = out_dir(getattr(args, "out", None), "clean")
    stem = safe_stem(src.stem)
    xlsx = unique_path(out / f"{stem}_clean.xlsx")
    _write_frame(df, xlsx)
    res.add_artifact(xlsx, "cleaned table (xlsx)")
    if getattr(args, "emit_csv", False):
        csvp = unique_path(out / f"{stem}_clean.csv")
        df.to_csv(csvp, index=False, encoding=CSV_ENCODING)
        res.add_artifact(csvp, "cleaned table (csv, UTF-8 BOM)")

    res.data.update(
        {
            "input": str(src.resolve()),
            "rows_before": int(original_shape[0]),
            "columns_before": int(original_shape[1]),
            "rows_after": int(len(df)),
            "columns_after": int(df.shape[1]),
            "steps": log,
            "final_columns": [str(c) for c in df.columns],
        }
    )
    return res


def _write_frame(df: pd.DataFrame, path: Path, *, sheet_name: str = "Sheet1") -> Path:
    """Write a frame to xlsx (or csv if the extension says so)."""
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    suf = path.suffix.lower()
    if suf == ".csv":
        df.to_csv(path, index=False, encoding=CSV_ENCODING)
        return path
    if suf == ".parquet":
        df.to_parquet(path, index=False)
        return path
    if suf == ".json":
        write_text(path, df.to_json(orient="records", force_ascii=False, indent=2))
        return path
    if suf == ".md":
        write_text(path, to_markdown(df))
        return path
    with pd.ExcelWriter(path, engine="xlsxwriter") as xw:
        df.to_excel(xw, index=False, sheet_name=sheet_name[:31])
        _autofit(xw, df, sheet_name[:31])
    return path


def _autofit(xw, df: pd.DataFrame, sheet: str, sample: int = 500) -> None:
    """Set sane column widths, accounting for wide CJK glyphs."""
    try:
        ws = xw.sheets[sheet]
    except Exception:  # noqa: BLE001
        return
    ws.freeze_panes(1, 0)
    ws.autofilter(0, 0, len(df), max(df.shape[1] - 1, 0))
    for idx, col in enumerate(df.columns):
        try:
            sample_vals = df[col].head(sample).astype("string").fillna("")
            width = max(
                [_display_width(str(col))]
                + [_display_width(v) for v in sample_vals.tolist()[:sample]]
                or [8]
            )
        except Exception:  # noqa: BLE001
            width = 12
        ws.set_column(idx, idx, min(max(width + 2, 8), 60))


def _display_width(text: str) -> int:
    return sum(2 if ord(ch) > 0x2E80 else 1 for ch in str(text))


def to_markdown(df: pd.DataFrame, max_rows: int = 200) -> str:
    """GitHub-flavoured markdown table without the optional `tabulate` dep."""
    d = df.head(max_rows)
    cols = [str(c) for c in d.columns]
    lines = ["| " + " | ".join(c.replace("|", "\\|") for c in cols) + " |"]
    lines.append("| " + " | ".join("---" for _ in cols) + " |")
    for rec in d.itertuples(index=False):
        cells = []
        for v in rec:
            s = "" if v is None or (isinstance(v, float) and math.isnan(v)) else str(v)
            cells.append(s.replace("|", "\\|").replace("\n", " "))
        lines.append("| " + " | ".join(cells) + " |")
    if len(df) > max_rows:
        lines.append(f"\n_({len(df) - max_rows} more rows omitted)_")
    return "\n".join(lines) + "\n"


# --------------------------------------------------------------------------
# profile
# --------------------------------------------------------------------------
def analyse_frame(df: pd.DataFrame, *, top_n: int = 10) -> dict[str, Any]:
    """The single analysis engine reused by profile/report/pivot/chart."""
    info: dict[str, Any] = {
        "rows": int(len(df)),
        "columns": int(df.shape[1]),
        "duplicate_rows": int(df.duplicated().sum()),
        "memory_bytes": int(df.memory_usage(deep=True).sum()),
        "column_details": [],
        "issues": [],
    }
    for c in df.columns:
        s = df[c]
        detail: dict[str, Any] = {
            "column": str(c),
            "dtype": str(s.dtype),
            "missing": int(s.isna().sum()),
            "missing_rate": round(float(s.isna().mean()), 4),
            "unique": int(s.nunique(dropna=True)),
        }
        non_null = s.dropna()
        if pd.api.types.is_numeric_dtype(s) and len(non_null):
            q = non_null.quantile([0.25, 0.5, 0.75])
            detail["kind"] = "numeric"
            detail["stats"] = {
                "min": _f(non_null.min()),
                "max": _f(non_null.max()),
                "mean": _f(non_null.mean()),
                "median": _f(non_null.median()),
                "std": _f(non_null.std()),
                "sum": _f(non_null.sum()),
                "q1": _f(q.loc[0.25]),
                "q3": _f(q.loc[0.75]),
                "skew": _f(non_null.skew()),
            }
            iqr = q.loc[0.75] - q.loc[0.25]
            if iqr > 0:
                out_mask = (non_null < q.loc[0.25] - 1.5 * iqr) | (non_null > q.loc[0.75] + 1.5 * iqr)
                detail["outliers"] = int(out_mask.sum())
        elif pd.api.types.is_datetime64_any_dtype(s) and len(non_null):
            detail["kind"] = "datetime"
            detail["stats"] = {
                "min": str(non_null.min()),
                "max": str(non_null.max()),
                "span_days": int((non_null.max() - non_null.min()).days),
            }
        else:
            detail["kind"] = "text"
            vc = non_null.astype("string").value_counts().head(top_n)
            detail["top_values"] = [{"value": str(k), "count": int(v)} for k, v in vc.items()]
            lens = non_null.astype("string").str.len()
            if len(lens):
                detail["text_length"] = {
                    "min": int(lens.min()),
                    "max": int(lens.max()),
                    "mean": round(float(lens.mean()), 1),
                }
            if detail["unique"] == len(df) and len(df) > 1:
                detail["likely_id"] = True
        if detail["missing_rate"] > 0.5 and len(df):
            info["issues"].append(
                {"column": str(c), "issue": "high_missing", "rate": detail["missing_rate"]}
            )
        if detail.get("likely_id"):
            info["issues"].append({"column": str(c), "issue": "all_unique_maybe_id"})
        if len(df) > 1 and detail["unique"] == 1:
            info["issues"].append({"column": str(c), "issue": "constant_column"})
        if 0.3 < detail["missing_rate"] <= 0.5 and len(df):
            info["issues"].append(
                {"column": str(c), "issue": "notable_missing", "rate": detail["missing_rate"]}
            )
        if detail["kind"] == "numeric" and detail.get("outliers"):
            info["issues"].append(
                {"column": str(c), "issue": "outliers_detected", "count": detail["outliers"]}
            )
        # Raw-data hygiene problems that a cleaning pass would fix: surface them
        # here too, otherwise `profile` on a dirty table looks falsely clean.
        if detail["kind"] == "text":
            raw = s.dropna().astype("string")
            if not raw.empty:
                n_full = int(raw.str.contains(r"[\uFF01-\uFF5E]", regex=True, na=False).sum())
                if n_full:
                    info["issues"].append(
                        {"column": str(c), "issue": "fullwidth_characters", "count": n_full}
                    )
                n_ws = int((raw.str.startswith(" ") | raw.str.endswith(" ")).sum())
                if n_ws:
                    info["issues"].append(
                        {"column": str(c), "issue": "untrimmed_whitespace", "count": n_ws}
                    )
                token_mask = raw.str.strip().str.lower().isin({t for t in NULL_TOKENS if t})
                n_token = int(token_mask.sum())
                if n_token:
                    info["issues"].append(
                        {"column": str(c), "issue": "null_like_tokens", "count": n_token,
                         "sample": raw[token_mask].head(3).tolist()}
                    )
                # Numeric-looking text stored as a string is a real type problem.
                if _looks_numeric(raw) and not raw.str.match(r"^-?\d+(\.\d+)?$", na=False).all():
                    info["issues"].append(
                        {"column": str(c), "issue": "numeric_stored_as_text",
                         "sample": raw.head(3).tolist()}
                    )
        info["column_details"].append(detail)

    num_cols = [str(c) for c in df.select_dtypes(include=[np.number]).columns]
    if len(num_cols) >= 2:
        try:
            corr = df[num_cols].corr(numeric_only=True)
            pairs = []
            for i, a in enumerate(num_cols):
                for b in num_cols[i + 1:]:
                    v = corr.loc[a, b]
                    if pd.notna(v) and abs(v) >= 0.5:
                        pairs.append({"a": a, "b": b, "r": round(float(v), 3)})
            pairs.sort(key=lambda x: -abs(x["r"]))
            info["strong_correlations"] = pairs[:20]
        except Exception:  # noqa: BLE001
            pass
    info["numeric_columns"] = num_cols
    info["datetime_columns"] = [str(c) for c in df.select_dtypes(include=["datetime", "datetimetz"]).columns]
    info["categorical_columns"] = [
        str(c) for c in df.columns
        if not pd.api.types.is_numeric_dtype(df[c])
        and not pd.api.types.is_datetime64_any_dtype(df[c])
        and df[c].nunique(dropna=True) <= max(50, int(len(df) * 0.5))
    ]
    return info


def _f(v: Any) -> Any:
    if v is None or (isinstance(v, float) and (math.isnan(v) or math.isinf(v))):
        return None
    try:
        f = float(v)
        return round(f, 6) if abs(f) < 1e15 else f
    except (TypeError, ValueError):
        return None


def cmd_profile(args) -> Result:
    res = Result("profile")
    src = resolve_inputs(args.input)[0]
    df = read_table(src, sheet=getattr(args, "sheet", None), header=getattr(args, "header", 0))
    info = analyse_frame(df, top_n=getattr(args, "top_n", 10))
    res.data["source"] = str(src.resolve())
    res.data["profile"] = info

    out = out_dir(getattr(args, "out", None), "profile")
    stem = safe_stem(src.stem)
    if getattr(args, "emit_markdown", True):
        md = profile_markdown(src.name, df, info)
        p = unique_path(out / f"{stem}_profile.md")
        write_text(p, md)
        res.add_artifact(p, "profile report (markdown)")
    if getattr(args, "emit_json", False):
        p = unique_path(out / f"{stem}_profile.json")
        write_text(p, json.dumps(info, ensure_ascii=False, indent=2, default=str))
        res.add_artifact(p, "profile (json)")
    return res


def profile_markdown(name: str, df: pd.DataFrame, info: dict[str, Any]) -> str:
    lines = [f"# 数据画像 / Data Profile — {name}", ""]
    lines.append(f"- 行数 Rows: **{info['rows']}**")
    lines.append(f"- 列数 Columns: **{info['columns']}**")
    lines.append(f"- 完全重复行 Duplicate rows: **{info['duplicate_rows']}**")
    lines.append(f"- 内存 Memory: {info['memory_bytes'] / 1024:.0f} KB")
    lines.append("")
    lines.append("## 字段概览 / Columns")
    lines.append("")
    lines.append("| 字段 | 类型 | 非空 | 缺失率 | 唯一值 | 关键统计 |")
    lines.append("| --- | --- | --- | --- | --- | --- |")
    for d in info["column_details"]:
        nonnull = info["rows"] - d["missing"]
        stat = ""
        if d["kind"] == "numeric" and d.get("stats"):
            s = d["stats"]
            stat = f"min={s['min']}, med={s['median']}, max={s['max']}, mean={s['mean']}"
            if d.get("outliers"):
                stat += f", 异常值={d['outliers']}"
        elif d["kind"] == "datetime" and d.get("stats"):
            stat = f"{d['stats']['min']} → {d['stats']['max']}"
        elif d.get("top_values"):
            stat = ", ".join(f"{v['value']}({v['count']})" for v in d["top_values"][:3])
        lines.append(
            f"| {d['column']} | {d['kind']} | {nonnull} | {d['missing_rate']:.1%} | {d['unique']} | {stat} |"
        )
    lines.append("")
    if info.get("issues"):
        lines.append("## 发现的问题 / Issues")
        lines.append("")
        for it in info["issues"]:
            extra = f" ({it['rate']:.1%})" if "rate" in it else ""
            lines.append(f"- `{it['column']}`: {it['issue']}{extra}")
        lines.append("")
    if info.get("strong_correlations"):
        lines.append("## 强相关字段 / Strong correlations (|r| ≥ 0.5)")
        lines.append("")
        lines.append("| A | B | r |")
        lines.append("| --- | --- | --- |")
        for p in info["strong_correlations"]:
            lines.append(f"| {p['a']} | {p['b']} | {p['r']} |")
        lines.append("")
    lines.append("## 数据预览 / Preview")
    lines.append("")
    lines.append(to_markdown(df.head(10)))
    return "\n".join(lines)


# --------------------------------------------------------------------------
# pivot / groupby
# --------------------------------------------------------------------------
_AGG_ALIASES = {
    "sum": "sum", "求和": "sum",
    "mean": "mean", "avg": "mean", "average": "mean", "平均": "mean",
    "count": "count", "计数": "count", "个数": "count",
    "min": "min", "最小": "min", "max": "max", "最大": "max",
    "median": "median", "中位数": "median",
    "std": "std", "stddev": "std", "标准差": "std",
    "nunique": "nunique", "去重计数": "nunique",
}


def _parse_aggs(spec: str | None) -> dict[str, str]:
    """'amount:sum,qty:mean' or 'sum' (applies to all value columns)."""
    if not spec:
        return {}
    out: dict[str, str] = {}
    for part in str(spec).split(","):
        part = part.strip()
        if not part:
            continue
        if ":" in part:
            col, agg = part.rsplit(":", 1)
            key = _AGG_ALIASES.get(agg.strip().lower(), agg.strip().lower())
            out[col.strip()] = key
        else:
            out["*"] = _AGG_ALIASES.get(part.lower(), part.lower())
    return out


def _split_csv_list(v: str | None) -> list[str]:
    return [x.strip() for x in str(v).split(",") if x.strip()] if v else []


def build_pivot(
    df: pd.DataFrame,
    index: list[str],
    columns: list[str],
    values: list[str],
    aggspec: dict[str, str],
    *,
    fill_value: Any = 0,
) -> tuple[pd.DataFrame, dict[str, str]]:
    missing = [c for c in index + columns + values if c not in df.columns]
    if missing:
        raise OfficeKitError(
            f"columns not found: {missing}; available: {[str(c) for c in df.columns]}"
        )
    work = df
    if not values:
        values = [c for c in df.columns if pd.api.types.is_numeric_dtype(df[c])][:5]
        if not values:
            values = [index[0]] if index else [df.columns[0]]
    default_agg = aggspec.get("*", "sum")
    agg_map: dict[str, Any] = {}
    simple = True
    for v in values:
        a = aggspec.get(v, default_agg)
        if not pd.api.types.is_numeric_dtype(work[v]) and a in ("sum", "mean", "median", "std"):
            a = "count"
        agg_map[v] = a
        if a != "sum":
            simple = False

    if len(values) == 1 and simple and not columns and len(index) == 1:
        tbl = (
            work.groupby(index, dropna=False)[values[0]]
            .agg(agg_map[values[0]])
            .reset_index()
            .sort_values(values[0], ascending=False)
        )
        return tbl, agg_map

    tbl = pd.pivot_table(
        work,
        index=index or None,
        columns=columns or None,
        values=values,
        aggfunc=agg_map,
        fill_value=fill_value,
        dropna=False,
        observed=False,
    )
    tbl = tbl.reset_index()
    tbl.columns = [
        "_".join(str(x) for x in col if str(x) not in ("", "nan")).strip("_")
        if isinstance(col, tuple) else str(col)
        for col in tbl.columns
    ]
    return tbl, agg_map


def cmd_pivot(args) -> Result:
    res = Result("pivot")
    src = resolve_inputs(args.input)[0]
    df = read_table(src, sheet=getattr(args, "sheet", None), header=getattr(args, "header", 0))
    index = _split_csv_list(getattr(args, "index", None))
    columns = _split_csv_list(getattr(args, "columns", None))
    values = _split_csv_list(getattr(args, "values", None))
    aggspec = _parse_aggs(getattr(args, "agg", None))
    if not index and not columns:
        # sensible default: first low-cardinality text column, then first numeric
        for c in df.columns:
            if not pd.api.types.is_numeric_dtype(df[c]) and 1 < df[c].nunique() <= 200:
                index = [str(c)]
                break
        if not index:
            index = [str(df.columns[0])]
        res.warn(f"no --index given; grouped by {index}")
    tbl, agg_map = build_pivot(df, index, columns, values, aggspec)
    res.data["source"] = str(src.resolve())
    res.data["index"] = index
    res.data["columns"] = columns
    res.data["values"] = values or "auto"
    res.data["agg"] = agg_map
    res.data["rows"] = int(len(tbl))
    res.data["preview"] = frame_preview(tbl, rows=15)

    out = out_dir(getattr(args, "out", None), "pivot")
    stem = safe_stem(getattr(args, "name", None) or src.stem)
    if getattr(args, "sort_by", None) and args.sort_by in tbl.columns:
        tbl = tbl.sort_values(args.sort_by, ascending=not getattr(args, "desc", False))
    xlsx = unique_path(out / f"{stem}_pivot.xlsx")
    with pd.ExcelWriter(xlsx, engine="xlsxwriter") as xw:
        tbl.to_excel(xw, index=False, sheet_name="pivot")
        _autofit(xw, tbl, "pivot")
    res.add_artifact(xlsx, "pivot table (xlsx)")
    return res


# --------------------------------------------------------------------------
# chart
# --------------------------------------------------------------------------
CHART_TYPES = ("bar", "barh", "line", "area", "pie", "scatter", "hist", "box", "heatmap")


def cmd_chart(args) -> Result:
    res = Result("chart")
    src = resolve_inputs(args.input)[0]
    df = read_table(src, sheet=getattr(args, "sheet", None), header=getattr(args, "header", 0))
    kind = (getattr(args, "type", None) or "auto").lower()

    plt, font = configure_matplotlib()
    if font is None:
        res.warn("no CJK font found; Chinese labels may render as boxes")

    x = getattr(args, "x", None)
    y = _split_csv_list(getattr(args, "y", None))
    group = getattr(args, "group", None)

    if kind == "auto":
        kind = _pick_chart(df, x, y, group)
        res.warn(f"chart type auto-selected: {kind}")
    if kind not in CHART_TYPES:
        raise OfficeKitError(f"unknown chart type {kind!r}; choose from {CHART_TYPES}")

    title = getattr(args, "title", None) or f"{src.stem} — {kind}"
    fig, ax = plt.subplots(figsize=_figsize(getattr(args, "figsize", None)))
    notes: list[str] = []

    if kind in ("bar", "barh", "line", "area"):
        x, y, agg = _resolve_xy(df, x, y, group, getattr(args, "agg", "sum"))
        if x is None:
            raise OfficeKitError("could not determine an x column; pass --x")
        top = getattr(args, "top", 30)
        plot_df = df
        if agg != "none" and group is None:
            plot_df = df.groupby(x, dropna=False)[y].agg(agg).reset_index()
        elif agg != "none" and group is not None:
            plot_df = df.groupby([x, group], dropna=False)[y].agg(agg).reset_index()
        if top and len(plot_df) > top:
            sort_col = y[0] if y[0] in plot_df.columns else plot_df.columns[-1]
            plot_df = plot_df.sort_values(sort_col, ascending=False).head(top)
            notes.append(f"showing top {top} of {len(df)} by {sort_col}")
        if group is not None and group in plot_df.columns:
            piv = plot_df.pivot_table(index=x, columns=group, values=y[0], aggfunc="first")
            _draw(piv, kind, ax, plt)
        else:
            _draw_frame(plot_df, x, y, kind, ax, plt)
        _rotate_ticks(ax, kind, plot_df, x, getattr(args, "rotate", None))

    elif kind == "pie":
        x, y, _ = _resolve_xy(df, x, y, None, getattr(args, "agg", "sum"))
        if x is None or not y:
            raise OfficeKitError("pie chart needs --x (label) and --y (value)")
        s = df.groupby(x, dropna=False)[y[0]].sum().sort_values(ascending=False)
        top = getattr(args, "top", 12)
        if len(s) > top:
            rest = s.iloc[top:].sum()
            s = pd.concat([s.iloc[:top], pd.Series({f"其他 ({len(s) - top})": rest})])
        ax.pie(s.values, labels=[str(i) for i in s.index], autopct="%1.1f%%",
               colors=PALETTE, startangle=90, textprops={"fontsize": 9})
        ax.axis("equal")

    elif kind == "scatter":
        x, y, _ = _resolve_xy(df, x, y, None, "none")
        if not x or not y:
            raise OfficeKitError("scatter needs --x and at least one --y")
        for i, col in enumerate(y):
            ax.scatter(df[x], df[col], s=18, alpha=0.65, color=PALETTE[i % len(PALETTE)], label=col)
        if len(y) > 1:
            ax.legend(fontsize=9)
        ax.set_xlabel(str(x))

    elif kind == "hist":
        cols = y or [str(c) for c in df.select_dtypes(include=[np.number]).columns][:4]
        if not cols:
            raise OfficeKitError("hist needs numeric columns; pass --y")
        ax.hist([df[c].dropna() for c in cols], bins=getattr(args, "bins", 20),
                label=[str(c) for c in cols], color=PALETTE[:len(cols)])
        if len(cols) > 1:
            ax.legend(fontsize=9)

    elif kind == "box":
        cols = y or [str(c) for c in df.select_dtypes(include=[np.number]).columns][:8]
        if not cols:
            raise OfficeKitError("box needs numeric columns; pass --y")
        ax.boxplot([df[c].dropna() for c in cols], tick_labels=[str(c) for c in cols])
        _rotate_ticks(ax, "bar", df, None, getattr(args, "rotate", 30))

    elif kind == "heatmap":
        cols = y or [str(c) for c in df.select_dtypes(include=[np.number]).columns]
        if len(cols) < 2:
            raise OfficeKitError("heatmap needs at least 2 numeric columns; pass --y")
        corr = df[cols].corr(numeric_only=True)
        im = ax.imshow(corr.values, cmap="RdBu_r", vmin=-1, vmax=1)
        ax.set_xticks(range(len(corr)), [str(c) for c in corr.columns], rotation=45, ha="right", fontsize=8)
        ax.set_yticks(range(len(corr)), [str(c) for c in corr.index], fontsize=8)
        for i in range(len(corr)):
            for j in range(len(corr)):
                v = corr.values[i, j]
                if pd.notna(v):
                    ax.text(j, i, f"{v:.2f}", ha="center", va="center", fontsize=7,
                            color="white" if abs(v) > 0.6 else "black")
        fig.colorbar(im, ax=ax, shrink=0.8)

    ax.set_title(str(title), fontsize=12, pad=12)
    if kind not in ("pie", "heatmap"):
        ax.grid(True, axis="y", alpha=0.25, linestyle="--")

    out = out_dir(getattr(args, "out", None), "chart")
    stem = safe_stem(getattr(args, "name", None) or f"{src.stem}_{kind}")
    png = unique_path(out / f"{stem}.png")
    fig.savefig(png, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    res.add_artifact(png, f"{kind} chart (png)")
    res.data.update({"type": kind, "title": str(title), "notes": notes})
    return res


def _figsize(spec: str | None):
    if spec:
        try:
            w, h = str(spec).lower().split("x")
            return (float(w), float(h))
        except Exception:  # noqa: BLE001
            pass
    return (11, 6)


def _pick_chart(df, x, y, group) -> str:
    num = [c for c in df.columns if pd.api.types.is_numeric_dtype(df[c])]
    if x and pd.api.types.is_datetime64_any_dtype(df[x]):
        return "line"
    if not x:
        if len(num) >= 2:
            return "heatmap"
        if num:
            return "hist"
        return "bar"
    if not pd.api.types.is_numeric_dtype(df[x]) and len(num) >= 1:
        if df[x].nunique() <= 12 and len(num) == 1:
            return "pie"
        return "bar"
    if len(num) >= 2:
        return "scatter"
    return "hist"


def _resolve_xy(df, x, y, group, default_agg):
    """Choose sensible x/y, aggregate only when it makes sense."""
    num = [str(c) for c in df.columns if pd.api.types.is_numeric_dtype(df[c])]
    cat = [str(c) for c in df.columns if not pd.api.types.is_numeric_dtype(df[c])]
    dt = [str(c) for c in df.select_dtypes(include=["datetime", "datetimetz"]).columns]
    agg = (default_agg or "sum").lower()
    if agg not in _AGG_ALIASES.values() and agg != "none":
        agg = _AGG_ALIASES.get(agg, "sum")
    if x is None:
        x = (dt + cat + num)[0] if (dt or cat or num) else None
    if not y:
        y = [c for c in num if c != x][:3] or ([c for c in cat if c != x][:1])
    # A pure scatter/hist needs raw rows, never aggregation: skip it when x is
    # effectively a row identifier.
    if x and len(df) == df[x].nunique():
        agg = "none"
    return x, y, agg


def _draw_frame(plot_df, x, y, kind, ax, plt):
    if kind == "line":
        for i, col in enumerate(y):
            ax.plot(plot_df[x], plot_df[col], marker="o", ms=3.5,
                    color=PALETTE[i % len(PALETTE)], label=col, linewidth=1.8)
    elif kind == "area":
        ax.stackplot(plot_df[x], *[plot_df[c] for c in y], labels=y,
                     colors=PALETTE[:len(y)], alpha=0.8)
    elif kind == "barh":
        ax.barh(plot_df[x].astype(str), plot_df[y[0]], color=PALETTE[0])
    else:
        width = 0.8 / max(len(y), 1)
        xs = np.arange(len(plot_df))
        for i, col in enumerate(y):
            ax.bar(xs + i * width, plot_df[col], width, label=col, color=PALETTE[i % len(PALETTE)])
        ax.set_xticks(xs + width * (len(y) - 1) / 2, plot_df[x].astype(str), rotation=30, ha="right", fontsize=8)
    if len(y) > 1:
        ax.legend(fontsize=9)
    if kind != "barh":
        ax.set_ylabel("值")


def _draw(piv, kind, ax, plt):
    if kind == "line":
        for i, col in enumerate(piv.columns):
            ax.plot(piv.index.astype(str), piv[col], marker="o", ms=3.5,
                    color=PALETTE[i % len(PALETTE)], label=str(col))
    elif kind == "area":
        ax.stackplot(piv.index.astype(str), *[piv[c] for c in piv.columns],
                     labels=[str(c) for c in piv.columns], colors=PALETTE[:piv.shape[1]], alpha=0.8)
    else:
        piv.plot(kind="bar", ax=ax, color=PALETTE[:piv.shape[1]], width=0.8)
    ax.legend(fontsize=8)


def _rotate_ticks(ax, kind, plot_df, x, rotate):
    if rotate is None:
        rotate = 30
    if kind == "barh" or x is None:
        return
    labels = [t.get_text() for t in ax.get_xticklabels()]
    if labels:
        ax.set_xticklabels(labels, rotation=rotate, ha="right", fontsize=8)


# --------------------------------------------------------------------------
# excel report with native charts
# --------------------------------------------------------------------------
def cmd_report(args) -> Result:
    res = Result("report")
    src = resolve_inputs(args.input)[0]
    df = read_table(src, sheet=getattr(args, "sheet", None), header=getattr(args, "header", 0))
    info = analyse_frame(df, top_n=getattr(args, "top_n", 10))
    group = getattr(args, "group", None)
    if not group:
        group = next(iter(info["categorical_columns"]), None)
    metrics = _split_csv_list(getattr(args, "metrics", None)) or info["numeric_columns"][:6]

    out = out_dir(getattr(args, "out", None), "report")
    stem = safe_stem(getattr(args, "name", None) or src.stem)
    title = getattr(args, "title", None) or f"{src.stem} 数据报告"

    xlsx = unique_path(out / f"{stem}_report.xlsx")
    sheet_summary, sheet_details = _build_xlsx_report(
        df, info, xlsx, title=title, group=group, metrics=metrics,
        top_n=getattr(args, "top_n", 10),
    )
    res.add_artifact(xlsx, "Excel report with charts and summary sheets")

    if getattr(args, "emit_markdown", True):
        md = _markdown_report(src.name, df, info, title=title, group=group, metrics=metrics)
        p = unique_path(out / f"{stem}_report.md")
        write_text(p, md)
        res.add_artifact(p, "same report as markdown")

    res.data.update(
        {
            "source": str(src.resolve()),
            "group_by": group,
            "metrics": metrics,
            "summary_columns": sheet_summary,
            "detail_sheets": sheet_details,
            "rows": int(len(df)),
        }
    )
    return res


def _build_xlsx_report(df, info, path, *, title, group, metrics, top_n) -> tuple[list[str], list[str]]:
    """Compose a formatted workbook: KPIs, stats, group breakdown + native charts."""
    wb_options = {"nan_inf_to_errors": True, "default_date_format": "yyyy-mm-dd"}
    wb = __import__("xlsxwriter").Workbook(str(path), wb_options)
    fmt = _report_formats(wb)

    ws = wb.add_worksheet("报告概览")
    ws.hide_gridlines(2)
    ws.set_column("A:A", 3)
    ws.set_column("B:B", 26)
    ws.set_column("C:H", 16)
    ws.write("B2", title, fmt["title"])
    ws.write("B3", f"数据源: {info.get('source_name', '')}    生成: {pd.Timestamp.now():%Y-%m-%d %H:%M}", fmt["sub"])
    row = 5
    ws.write_row(row, 1, ["行数", "列数", "重复行", "数值字段", "文本字段", "日期字段"], fmt["hdr"])
    row += 1
    ws.write_row(row, 1, [info["rows"], info["columns"], info["duplicate_rows"],
                          len(info["numeric_columns"]), len(info["categorical_columns"]),
                          len(info["datetime_columns"])], fmt["kpi"])
    row += 2

    ws.write(row, 1, "字段统计", fmt["h2"])
    row += 1
    ws.write_row(row, 1, ["字段", "类型", "非空", "缺失率", "唯一值", "最小", "中位数", "最大", "均值", "异常值"], fmt["hdr"])
    row += 1
    for d in info["column_details"]:
        s = d.get("stats", {}) or {}
        ws.write_row(row, 1, [
            d["column"], d["kind"], info["rows"] - d["missing"],
            d["missing_rate"], d["unique"],
            _num_or_blank(s.get("min")), _num_or_blank(s.get("median")),
            _num_or_blank(s.get("max")), _num_or_blank(s.get("mean")),
            d.get("outliers", ""),
        ], fmt["cell"])
        ws.set_column(3, 3, 12, fmt["pct"])
        row += 1
    row += 1

    chart_anchor_row = row
    if group and group in df.columns and metrics:
        agg_col = metrics[0]
        try:
            g = (df.groupby(group, dropna=False)[agg_col]
                 .agg(["sum", "mean", "count"])
                 .reset_index()
                 .sort_values("sum", ascending=False)
                 .head(top_n))
            g.columns = [group, f"{agg_col}_合计", f"{agg_col}_平均", "记录数"]
            ws.write(row, 1, f"按 {group} 汇总（{agg_col}）", fmt["h2"])
            row += 1
            start = row
            ws.write_row(row, 1, [str(c) for c in g.columns], fmt["hdr"])
            row += 1
            for rec in g.itertuples(index=False):
                ws.write_row(row, 1, list(rec), fmt["cell"])
                row += 1

            ch = wb.add_chart({"type": "column"})
            ch.add_series({
                "name": f"{agg_col} 合计",
                "categories": ["报告概览", start, 1, start + len(g) - 1, 1],
                "values": ["报告概览", start, 2, start + len(g) - 1, 2],
                "fill": {"color": PALETTE[0]},
            })
            ch.set_title({"name": f"各 {group} 的 {agg_col} 合计"})
            ch.set_legend({"none": True})
            ch.set_size({"width": 620, "height": 340})
            ws.insert_chart(chart_anchor_row - 1, 8, ch)
        except Exception as exc:  # noqa: BLE001
            ws.write(row, 1, f"汇总失败: {exc}", fmt["cell"])
            row += 1

    detail_sheets: list[str] = []
    ws_d = wb.add_worksheet("明细数据")
    ws_d.write_row(0, 0, [str(c) for c in df.columns], fmt["hdr"])
    for i, rec in enumerate(df.itertuples(index=False), start=1):
        ws_d.write_row(i, 0, [_xl_safe(v) for v in rec], fmt["cell"])
    ws_d.freeze_panes(1, 0)
    ws_d.autofilter(0, 0, len(df), max(df.shape[1] - 1, 0))
    for idx, col in enumerate(df.columns):
        try:
            width = max([_display_width(str(col))] + [_display_width(str(v)) for v in df[col].head(300)])
        except Exception:  # noqa: BLE001
            width = 12
        ws_d.set_column(idx, idx, min(max(width + 2, 8), 50))
    detail_sheets.append("明细数据")

    if metrics:
        ws_n = wb.add_worksheet("数值分布")
        ws_n.write_row(0, 0, ["字段", "最小", "P25", "中位数", "P75", "最大", "均值", "标准差", "偏度"], fmt["hdr"])
        r = 1
        for d in info["column_details"]:
            if d["kind"] != "numeric" or not d.get("stats"):
                continue
            s = d["stats"]
            ws_n.write_row(r, 0, [d["column"], s.get("min"), s.get("q1"), s.get("median"),
                                  s.get("q3"), s.get("max"), s.get("mean"), s.get("std"), s.get("skew")], fmt["cell"])
            r += 1
        ws_n.set_column(0, 0, 24)
        ws_n.set_column(1, 8, 12)
        detail_sheets.append("数值分布")
        if r > 1:
            hist = wb.add_chart({"type": "column"})
            hist.add_series({
                "name": "均值",
                "categories": ["数值分布", 1, 0, r - 1, 0],
                "values": ["数值分布", 1, 6, r - 1, 6],
                "fill": {"color": PALETTE[2]},
            })
            hist.set_title({"name": "各数值字段均值对比"})
            hist.set_legend({"none": True})
            hist.set_size({"width": 620, "height": 340})
            ws_n.insert_chart(1, 10, hist)

    wb.close()
    return [f"字段: {d['column']}" for d in info["column_details"]], detail_sheets


def _num_or_blank(v):
    return "" if v is None else v


def _xl_safe(v):
    """xlsxwriter refuses some objects; coerce to something writable."""
    if v is None:
        return ""
    if isinstance(v, (int, float, str, bool)):
        if isinstance(v, float) and (math.isnan(v) or math.isinf(v)):
            return ""
        return v
    if hasattr(v, "isoformat"):
        return str(v)
    if isinstance(v, (list, tuple, dict, set)):
        return json.dumps(v, ensure_ascii=False, default=str)
    return str(v)


def _report_formats(wb) -> dict[str, Any]:
    return {
        "title": wb.add_format({"font_size": 18, "bold": True, "font_color": "#1F3864"}),
        "sub": wb.add_format({"font_size": 10, "font_color": "#808080"}),
        "h2": wb.add_format({"font_size": 13, "bold": True, "font_color": "#1F3864",
                             "bottom": 2, "bottom_color": "#1F3864"}),
        "hdr": wb.add_format({"bold": True, "bg_color": "#1F3864", "font_color": "white",
                              "border": 1, "border_color": "#BFBFBF", "align": "center",
                              "valign": "vcenter", "text_wrap": True}),
        "kpi": wb.add_format({"font_size": 14, "bold": True, "align": "center",
                              "bg_color": "#EDF2FA", "border": 1, "border_color": "#BFBFBF"}),
        "cell": wb.add_format({"border": 1, "border_color": "#D9D9D9", "valign": "top"}),
        "pct": wb.add_format({"num_format": "0.0%", "border": 1, "border_color": "#D9D9D9"}),
    }


def _markdown_report(name, df, info, *, title, group, metrics) -> str:
    L = [f"# {title}", "", f"数据源：`{name}` ・ 行数 **{info['rows']}** ・ 列数 **{info['columns']}**", ""]
    L.append("## 1. 关键指标")
    L.append("")
    L.append(f"- 完全重复行：**{info['duplicate_rows']}**")
    L.append(f"- 数值字段：{', '.join(info['numeric_columns']) or '无'}")
    L.append(f"- 日期字段：{', '.join(info['datetime_columns']) or '无'}")
    L.append("")
    L.append("## 2. 字段统计")
    L.append("")
    L.append("| 字段 | 类型 | 缺失率 | 唯一值 | 统计 |")
    L.append("| --- | --- | --- | --- | --- |")
    for d in info["column_details"]:
        s = d.get("stats", {}) or {}
        stat = ""
        if d["kind"] == "numeric":
            stat = f"min={s.get('min')}, 中位={s.get('median')}, max={s.get('max')}, 均值={s.get('mean')}"
        elif d["kind"] == "datetime":
            stat = f"{s.get('min')} → {s.get('max')}"
        elif d.get("top_values"):
            stat = "、".join(f"{v['value']}({v['count']})" for v in d["top_values"][:3])
        L.append(f"| {d['column']} | {d['kind']} | {d['missing_rate']:.1%} | {d['unique']} | {stat} |")
    L.append("")
    if group and group in df.columns and metrics:
        agg = metrics[0]
        L.append(f"## 3. 按 `{group}` 汇总（{agg}）")
        L.append("")
        try:
            g = (df.groupby(group, dropna=False)[agg].agg(["sum", "mean", "count"])
                 .sort_values("sum", ascending=False).head(20).reset_index())
            g.columns = [group, f"{agg}_合计", f"{agg}_平均", "记录数"]
            L.append(to_markdown(g))
        except Exception as exc:  # noqa: BLE001
            L.append(f"_汇总失败: {exc}_")
        L.append("")
    if info.get("strong_correlations"):
        L.append("## 4. 强相关字段")
        L.append("")
        L.append("| 字段 A | 字段 B | 相关系数 |")
        L.append("| --- | --- | --- |")
        for p in info["strong_correlations"]:
            L.append(f"| {p['a']} | {p['b']} | {p['r']} |")
        L.append("")
    if info.get("issues"):
        L.append("## 5. 数据质量问题")
        L.append("")
        for it in info["issues"]:
            extra = f"（{it['rate']:.1%}）" if "rate" in it else ""
            L.append(f"- `{it['column']}`：{it['issue']}{extra}")
        L.append("")
    L.append("## 6. 数据预览")
    L.append("")
    L.append(to_markdown(df.head(15)))
    return "\n".join(L)


# --------------------------------------------------------------------------
# merge / compare / convert
# --------------------------------------------------------------------------
def cmd_merge(args) -> Result:
    res = Result("merge")
    files = resolve_inputs(args.input)
    frames: list[pd.DataFrame] = []
    sources: list[dict[str, Any]] = []
    for f in files:
        df = read_table(f, sheet=getattr(args, "sheet", None), header=getattr(args, "header", 0))
        if getattr(args, "add_source", True):
            df = df.copy()
            df.insert(0, "_source_file", f.name)
        frames.append(df)
        sources.append({"file": f.name, "rows": int(len(df)), "columns": int(df.shape[1])})

    mode = (getattr(args, "mode", None) or "auto").lower()
    if mode == "auto":
        shapes = {tuple(str(c) for c in fr.columns) for fr in frames}
        mode = "concat" if len(shapes) == 1 else "align"
    if mode == "concat":
        merged = pd.concat(frames, ignore_index=True, sort=False)
    elif mode == "align":
        merged = pd.concat(frames, ignore_index=True, sort=False)
        res.warn("column sets differ; aligned on the union of columns (missing filled with NaN)")
    elif mode == "join":
        key = getattr(args, "key", None)
        if not key:
            raise OfficeKitError("mode=join requires --key")
        keys = _split_csv_list(key)
        merged = frames[0]
        for nxt in frames[1:]:
            missing = [k for k in keys if k not in nxt.columns]
            if missing:
                raise OfficeKitError(f"join key {missing} missing in the next file")
            merged = merged.merge(nxt, on=keys, how=getattr(args, "how", "left"),
                                  suffixes=("", f"_{len(sources)}"))
    else:
        raise OfficeKitError("mode must be auto, concat, align or join")

    if getattr(args, "dedupe", False):
        before = len(merged)
        key_cols = _split_csv_list(getattr(args, "dedupe_keys", None)) or None
        merged = merged.drop_duplicates(subset=key_cols)
        res.data["dedupe_removed"] = int(before - len(merged))

    out = out_dir(getattr(args, "out", None), "merge")
    stem = safe_stem(getattr(args, "name", None) or "merged")
    xlsx = unique_path(out / f"{stem}.xlsx")
    _write_frame(merged, xlsx)
    res.add_artifact(xlsx, "merged table (xlsx)")
    res.data.update({"mode": mode, "sources": sources, "rows": int(len(merged)),
                     "columns": int(merged.shape[1])})
    return res


def cmd_compare(args) -> Result:
    res = Result("compare")
    files = resolve_inputs(args.input)
    if len(files) < 2:
        raise OfficeKitError("compare needs two files (or two --key groups)")
    a = read_table(files[0], sheet=getattr(args, "sheet", None))
    b = read_table(files[1], sheet=getattr(args, "sheet", None))
    keys = _split_csv_list(getattr(args, "key", None)) or _common_key(a, b)

    report: dict[str, Any] = {
        "a": {"file": files[0].name, "rows": int(len(a)), "columns": int(a.shape[1])},
        "b": {"file": files[1].name, "rows": int(len(b)), "columns": int(b.shape[1])},
        "key": keys,
        "only_in_a_columns": [str(c) for c in a.columns if c not in b.columns],
        "only_in_b_columns": [str(c) for c in b.columns if c not in a.columns],
    }
    out = out_dir(getattr(args, "out", None), "compare")

    if keys:
        ka = set(map(tuple, a[keys].astype(str).to_numpy()))
        kb = set(map(tuple, b[keys].astype(str).to_numpy()))
        only_a, only_b = sorted(ka - kb), sorted(kb - ka)
        report["rows_only_in_a"] = len(only_a)
        report["rows_only_in_b"] = len(only_b)
        report["sample_only_in_a"] = [dict(zip(keys, r)) for r in only_a[:20]]
        report["sample_only_in_b"] = [dict(zip(keys, r)) for r in only_b[:20]]
        if only_a:
            p = unique_path(out / "rows_only_in_A.csv")
            pd.DataFrame(only_a, columns=keys).to_csv(p, index=False, encoding=CSV_ENCODING)
            res.add_artifact(p, "keys present only in file A")
        if only_b:
            p = unique_path(out / "rows_only_in_B.csv")
            pd.DataFrame(only_b, columns=keys).to_csv(p, index=False, encoding=CSV_ENCODING)
            res.add_artifact(p, "keys present only in file B")

        common_cols = [c for c in a.columns if c in b.columns and c not in keys]
        changes = []
        try:
            m = a.merge(b, on=keys, suffixes=("_a", "_b"), how="inner")
            for c in common_cols:
                ca, cb = f"{c}_a", f"{c}_b"
                if ca not in m.columns or cb not in m.columns:
                    continue
                diff = ~(
                    (m[ca].astype("string").fillna("") == m[cb].astype("string").fillna(""))
                )
                if diff.any():
                    changes.append(
                        {
                            "column": str(c),
                            "changed_rows": int(diff.sum()),
                            "sample": [
                                {"key": dict(zip(keys, row[keys])), "a": _xl_safe(row[ca]), "b": _xl_safe(row[cb])}
                                for _, row in m[diff].head(10).iterrows()
                            ],
                        }
                    )
        except Exception as exc:  # noqa: BLE001
            res.warn(f"value comparison skipped: {exc}")
        report["changed_columns"] = changes
    else:
        res.warn("no common key column found; compared shapes and column names only")

    p = unique_path(out / "compare.json")
    write_text(p, json.dumps(report, ensure_ascii=False, indent=2, default=_xl_safe))
    res.add_artifact(p, "comparison detail (json)")
    res.data.update(report)
    return res


def _common_key(a: pd.DataFrame, b: pd.DataFrame) -> list[str]:
    shared = [c for c in a.columns if c in b.columns]
    for c in shared:
        if a[c].nunique() == len(a) and b[c].nunique() == len(b):
            return [str(c)]
    for c in shared:
        if not pd.api.types.is_numeric_dtype(a[c]) and a[c].nunique() > 1:
            return [str(c)]
    return []


def cmd_convert(args) -> Result:
    res = Result("convert")
    files = resolve_inputs(args.input)
    to = (getattr(args, "to", None) or "").lower().lstrip(".")
    if not to:
        raise OfficeKitError("--to is required (xlsx, csv, json, md, parquet, tsv)")
    out = out_dir(getattr(args, "out", None), "convert")
    produced: list[dict[str, Any]] = []

    for f in files:
        df = read_table(f, sheet=getattr(args, "sheet", None), header=getattr(args, "header", 0))
        target = unique_path(out / f"{safe_stem(f.stem)}.{to}")
        _write_frame(df, target)
        res.add_artifact(target, f"{f.name} -> {to}")
        produced.append({"source": f.name, "target": target.name, "rows": int(len(df))})

    if to == "md" and getattr(args, "combine", False) and len(files) > 1:
        combined = out / "all_tables.md"
        parts = [f"# {f.name}\n\n" + to_markdown(read_table(f)) for f in files]
        write_text(combined, "\n\n".join(parts))
        res.add_artifact(combined, "all tables in one markdown file")

    res.data.update({"to": to, "converted": produced})
    return res
