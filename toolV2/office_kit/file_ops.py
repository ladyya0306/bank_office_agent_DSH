"""File organization: inventory, dedupe, batch rename, rule-based归类.

Destructive operations are dry-run by default: the caller must pass --apply, and
every plan is written to a CSV/JSON log first so a move can be undone.
"""
from __future__ import annotations

import csv
import fnmatch
import hashlib
import json
import os
import re
import shutil
from collections import defaultdict
from datetime import datetime
from pathlib import Path
from typing import Any

from .common import (
    OfficeKitError,
    Result,
    ensure_parent,
    human_bytes,
    out_dir,
    resolve_roots,
    safe_stem,
    unique_path,
    write_text,
)

# Directories that are almost never the user's real documents.
JUNK_DIRS = {
    "$recycle.bin", "system volume information", "node_modules", ".git", ".svn",
    "__pycache__", ".pytest_cache", ".mypy_cache", ".ruff_cache", ".venv", "venv",
    "env", ".idea", ".vscode", "dist", "build", ".next", ".cache", "target",
    "site-packages", "appdata",
}

DEFAULT_EXCLUDES = ["~$*", "*.tmp", "*.temp", ".ds_store", "thumbs.db", "desktop.ini"]

CATEGORY_MAP: dict[str, set[str]] = {
    "文档_Documents": {".doc", ".docx", ".pdf", ".txt", ".rtf", ".odt", ".md", ".wps", ".pages"},
    "表格_Spreadsheets": {".xls", ".xlsx", ".xlsm", ".csv", ".ods", ".et"},
    "演示_Presentations": {".ppt", ".pptx", ".odp", ".dps", ".key"},
    "图片_Images": {".jpg", ".jpeg", ".png", ".gif", ".bmp", ".tif", ".tiff", ".webp", ".svg", ".heic"},
    "音频_Audio": {".mp3", ".wav", ".flac", ".aac", ".m4a", ".wma", ".ogg"},
    "视频_Video": {".mp4", ".avi", ".mkv", ".mov", ".wmv", ".flv", ".webm", ".m4v"},
    "压缩包_Archives": {".zip", ".rar", ".7z", ".tar", ".gz", ".bz2", ".xz"},
    "程序_Code": {".py", ".js", ".ts", ".java", ".c", ".cpp", ".cs", ".go", ".rs", ".php", ".rb", ".sh", ".ps1", ".sql", ".html", ".css", ".json", ".xml", ".yml", ".yaml"},
    "可执行_Executables": {".exe", ".msi", ".bat", ".cmd", ".app", ".dmg", ".apk"},
    "字体_Fonts": {".ttf", ".otf", ".woff", ".woff2", ".fon"},
    "电子书_eBooks": {".epub", ".mobi", ".azw3", ".djvu"},
}


def _suffix_category(suffix: str) -> str:
    s = suffix.lower()
    for cat, exts in CATEGORY_MAP.items():
        if s in exts:
            return cat
    return "其他_Other"


# --------------------------------------------------------------------------
# traversal
# --------------------------------------------------------------------------
def walk_files(
    root: Path,
    *,
    recursive: bool = True,
    include: list[str] | None = None,
    exclude: list[str] | None = None,
    skip_junk: bool = True,
    max_files: int = 500000,
) -> list[Path]:
    root = Path(root)
    if not root.exists():
        raise OfficeKitError(f"directory not found: {root}")
    include = include or []
    exclude = list(exclude or []) + DEFAULT_EXCLUDES
    found: list[Path] = []

    def ok(p: Path) -> bool:
        name = p.name.lower()
        if any(fnmatch.fnmatch(name, pat.lower()) for pat in exclude):
            return False
        if include and not any(fnmatch.fnmatch(name, pat.lower()) for pat in include):
            return False
        return True

    if recursive:
        for dirpath, dirnames, filenames in os.walk(root):
            if skip_junk:
                dirnames[:] = [d for d in dirnames if d.lower() not in JUNK_DIRS]
            for fn in filenames:
                p = Path(dirpath) / fn
                if ok(p):
                    found.append(p)
                    if len(found) >= max_files:
                        return found
    else:
        for p in sorted(root.iterdir()):
            if p.is_file() and ok(p):
                found.append(p)
    return found


def file_info(p: Path) -> dict[str, Any]:
    try:
        st = p.stat()
    except OSError as exc:
        return {"path": str(p), "error": str(exc)}
    return {
        "path": str(p),
        "name": p.name,
        "suffix": p.suffix.lower(),
        "category": _suffix_category(p.suffix),
        "bytes": st.st_size,
        "modified": datetime.fromtimestamp(st.st_mtime).isoformat(sep=" ", timespec="seconds"),
        "created": datetime.fromtimestamp(st.st_ctime).isoformat(sep=" ", timespec="seconds"),
        "age_days": round((datetime.now() - datetime.fromtimestamp(st.st_mtime)).days, 1),
    }


def digest(path: Path, *, algo: str = "sha256", chunk: int = 1 << 20) -> str:
    """Hash by a cheap fingerprint first: full hash only for same-size candidates."""
    h = hashlib.new(algo)
    with open(path, "rb") as fh:
        while True:
            block = fh.read(chunk)
            if not block:
                break
            h.update(block)
    return h.hexdigest()


def quick_key(path: Path) -> str:
    """size + first/last 64 KB: cheap pre-filter before full hashing."""
    size = path.stat().st_size
    h = hashlib.sha256()
    with open(path, "rb") as fh:
        h.update(fh.read(65536))
        if size > 131072:
            fh.seek(-65536, os.SEEK_END)
            h.update(fh.read(65536))
    return f"{size}:{h.hexdigest()}"


# --------------------------------------------------------------------------
# command: dir list
# --------------------------------------------------------------------------
def cmd_list(args) -> Result:
    res = Result("list")
    roots = resolve_roots(args.input)
    recursive = not getattr(args, "shallow", False)
    include = _split(getattr(args, "include", None))
    exclude = _split(getattr(args, "exclude", None))
    files = []
    for root in roots:
        if root.is_file():
            files.append(root)
        else:
            files.extend(walk_files(root, recursive=recursive, include=include, exclude=exclude))
    infos = [file_info(p) for p in files]
    infos.sort(key=lambda d: d.get("modified", ""), reverse=True)
    res.data.update(
        {
            "roots": [str(r) for r in roots],
            "count": len(infos),
            "total_bytes": sum(i.get("bytes", 0) for i in infos),
            "total_human": human_bytes(sum(i.get("bytes", 0) for i in infos)),
            "files": infos[: int(getattr(args, "limit", 200) or 200)],
            "truncated": len(infos) > int(getattr(args, "limit", 200) or 200),
        }
    )
    return res


def _split(v: str | None) -> list[str]:
    return [x.strip() for x in str(v).split(",") if x.strip()] if v else []


# --------------------------------------------------------------------------
# command: dir report
# --------------------------------------------------------------------------
def cmd_report(args) -> Result:
    res = Result("dir-report")
    roots = resolve_roots(args.input)
    recursive = not getattr(args, "shallow", False)
    files: list[Path] = []
    for root in roots:
        if root.is_file():
            files.append(root)
        else:
            files.extend(walk_files(root, recursive=recursive,
                                    include=_split(getattr(args, "include", None)),
                                    exclude=_split(getattr(args, "exclude", None))))
    if not files:
        raise OfficeKitError("no files found under the given paths")

    infos = [file_info(p) for p in files]
    by_cat: dict[str, dict[str, Any]] = defaultdict(lambda: {"count": 0, "bytes": 0, "extensions": defaultdict(int)})
    by_ext: dict[str, dict[str, Any]] = defaultdict(lambda: {"count": 0, "bytes": 0})
    for i in infos:
        cat = i["category"]
        by_cat[cat]["count"] += 1
        by_cat[cat]["bytes"] += i["bytes"]
        by_cat[cat]["extensions"][i["suffix"] or "(none)"] += 1
        by_ext[i["suffix"] or "(none)"]["count"] += 1
        by_ext[i["suffix"] or "(none)"]["bytes"] += i["bytes"]

    # duplicates: only hash files whose (size, edge bytes) fingerprint collides
    dup_groups: list[dict[str, Any]] = []
    if getattr(args, "duplicates", True):
        buckets: dict[str, list[Path]] = defaultdict(list)
        for i in infos:
            if i.get("bytes", 0) == 0:
                continue
            try:
                buckets[quick_key(Path(i["path"]))].append(Path(i["path"]))
            except OSError:
                continue
        for key, group in buckets.items():
            if len(group) < 2:
                continue
            by_hash: dict[str, list[Path]] = defaultdict(list)
            for p in group:
                try:
                    by_hash[digest(p)].append(p)
                except OSError:
                    continue
            for h, same in by_hash.items():
                if len(same) > 1:
                    waste = same[0].stat().st_size * (len(same) - 1)
                    dup_groups.append(
                        {
                            "hash": h[:16],
                            "count": len(same),
                            "bytes_each": same[0].stat().st_size,
                            "wasted_bytes": waste,
                            "files": [str(p) for p in same],
                        }
                    )
        dup_groups.sort(key=lambda g: -g["wasted_bytes"])

    empty_files = [i["path"] for i in infos if i.get("bytes") == 0]
    empty_dirs: list[str] = []
    for root in roots:
        if root.is_dir():
            for dirpath, dirnames, filenames in os.walk(root):
                dirnames[:] = [d for d in dirnames if d.lower() not in JUNK_DIRS]
                if not dirnames and not filenames and Path(dirpath) != root:
                    empty_dirs.append(dirpath)

    oldest = sorted(infos, key=lambda i: i.get("modified", ""))[:10]
    newest = sorted(infos, key=lambda i: i.get("modified", ""), reverse=True)[:10]
    largest = sorted(infos, key=lambda i: -i.get("bytes", 0))[:20]
    stale = [i for i in infos if i.get("age_days", 0) > float(getattr(args, "stale_days", 365) or 365)]

    summary = {
        "roots": [str(r) for r in roots],
        "file_count": len(infos),
        "total_bytes": sum(i.get("bytes", 0) for i in infos),
        "total_human": human_bytes(sum(i.get("bytes", 0) for i in infos)),
        "by_category": {
            k: {"count": v["count"], "bytes": v["bytes"], "human": human_bytes(v["bytes"]),
                "top_extensions": dict(sorted(v["extensions"].items(), key=lambda kv: -kv[1])[:5])}
            for k, v in sorted(by_cat.items(), key=lambda kv: -kv[1]["bytes"])
        },
        "by_extension": {
            k: {"count": v["count"], "bytes": v["bytes"], "human": human_bytes(v["bytes"])}
            for k, v in sorted(by_ext.items(), key=lambda kv: -kv[1]["bytes"])[:30]
        },
        "duplicate_groups": len(dup_groups),
        "duplicate_wasted_bytes": sum(g["wasted_bytes"] for g in dup_groups),
        "duplicate_wasted_human": human_bytes(sum(g["wasted_bytes"] for g in dup_groups)),
        "empty_files": len(empty_files),
        "empty_dirs_count": len(empty_dirs),
        "stale_files_over_days": float(getattr(args, "stale_days", 365) or 365),
        "stale_count": len(stale),
        "stale_bytes": sum(i.get("bytes", 0) for i in stale),
        "largest": largest[:10],
        "oldest": oldest[:5],
        "newest": newest[:5],
    }
    res.data["summary"] = summary

    out = out_dir(getattr(args, "out", None), "dir-report")
    stem = safe_stem(Path(roots[0]).name or "files")
    if getattr(args, "emit_json", True):
        p = unique_path(out / f"{stem}_inventory.json")
        write_text(p, json.dumps({"summary": summary, "duplicates": dup_groups,
                                  "empty_dirs": empty_dirs, "empty_files": empty_files,
                                  "files": infos}, ensure_ascii=False, indent=2, default=str))
        res.add_artifact(p, "full inventory (json)")
    if getattr(args, "emit_markdown", True):
        p = unique_path(out / f"{stem}_inventory.md")
        write_text(p, _inventory_markdown(summary, dup_groups, empty_dirs, infos))
        res.add_artifact(p, "human-readable inventory report")
    if dup_groups and getattr(args, "emit_csv", True):
        p = unique_path(out / f"{stem}_duplicates.csv")
        ensure_parent(p)
        with open(p, "w", encoding="utf-8-sig", newline="") as fh:
            w = csv.writer(fh)
            w.writerow(["group_hash", "group_index", "keep_or_dup", "path", "bytes"])
            for g in dup_groups:
                for idx, f in enumerate(g["files"]):
                    w.writerow([g["hash"], idx, "keep" if idx == 0 else "duplicate", f, g["bytes_each"]])
        res.add_artifact(p, "duplicate files (csv)")
    return res


def _inventory_markdown(summary, dup_groups, empty_dirs, infos) -> str:
    L = ["# 文件清单报告 / Directory Inventory", ""]
    L.append(f"- 文件总数 Files: **{summary['file_count']}**")
    L.append(f"- 占用空间 Total size: **{summary['total_human']}**")
    L.append(f"- 重复文件组 Duplicate groups: **{summary['duplicate_groups']}** （可回收 {summary['duplicate_wasted_human']}）")
    L.append(f"- 空文件 Empty files: **{summary['empty_files']}** ・ 空目录 Empty dirs: **{summary['empty_dirs_count']}**")
    L.append(f"- 超过 {summary['stale_files_over_days']:.0f} 天未修改: **{summary['stale_count']}** 个，共 {human_bytes(summary['stale_bytes'])}")
    L.append("")
    L.append("## 按类别 / By category")
    L.append("")
    L.append("| 类别 | 文件数 | 大小 | 常见扩展名 |")
    L.append("| --- | --- | --- | --- |")
    for k, v in summary["by_category"].items():
        exts = ", ".join(f"{e}({c})" for e, c in v["top_extensions"].items())
        L.append(f"| {k} | {v['count']} | {v['human']} | {exts} |")
    L.append("")
    L.append("## 按扩展名 Top / By extension")
    L.append("")
    L.append("| 扩展名 | 文件数 | 大小 |")
    L.append("| --- | --- | --- |")
    for k, v in list(summary["by_extension"].items())[:20]:
        L.append(f"| {k} | {v['count']} | {v['human']} |")
    L.append("")
    if summary["largest"]:
        L.append("## 最大的文件 / Largest files")
        L.append("")
        L.append("| 文件 | 大小 | 修改时间 |")
        L.append("| --- | --- | --- |")
        for i in summary["largest"]:
            L.append(f"| {i['name']} | {human_bytes(i['bytes'])} | {i['modified']} |")
        L.append("")
    if dup_groups:
        L.append(f"## 重复文件 / Duplicates (top {min(len(dup_groups), 15)})")
        L.append("")
        for g in dup_groups[:15]:
            L.append(f"- **{human_bytes(g['bytes_each'])} × {g['count']}** 可回收 {human_bytes(g['wasted_bytes'])}")
            for idx, f in enumerate(g["files"]):
                L.append(f"  - {'保留' if idx == 0 else '重复'} `{f}`")
        L.append("")
    if empty_dirs:
        L.append("## 空目录 / Empty directories")
        L.append("")
        for d in empty_dirs[:30]:
            L.append(f"- `{d}`")
        L.append("")
    return "\n".join(L)


# --------------------------------------------------------------------------
# command: dedupe
# --------------------------------------------------------------------------
def cmd_dedupe(args) -> Result:
    res = Result("dedupe")
    roots = resolve_roots(args.input)
    files: list[Path] = []
    for root in roots:
        files.extend(walk_files(root, recursive=not getattr(args, "shallow", False),
                                include=_split(getattr(args, "include", None)),
                                exclude=_split(getattr(args, "exclude", None))))
    buckets: dict[str, list[Path]] = defaultdict(list)
    for p in files:
        try:
            if p.stat().st_size == 0:
                continue
            buckets[quick_key(p)].append(p)
        except OSError:
            continue

    groups: list[dict[str, Any]] = []
    for group in buckets.values():
        if len(group) < 2:
            continue
        by_hash: dict[str, list[Path]] = defaultdict(list)
        for p in group:
            try:
                by_hash[digest(p)].append(p)
            except OSError:
                continue
        for h, same in by_hash.items():
            if len(same) > 1:
                # keep the oldest (usually the original) unless told otherwise
                same.sort(key=lambda p: p.stat().st_mtime)
                groups.append({"hash": h, "files": same, "bytes_each": same[0].stat().st_size})
    if not groups:
        res.data.update(
            {
                "duplicate_groups": 0,
                "duplicate_files": 0,
                "reclaimable_bytes": 0,
                "reclaimable_human": "0 B",
                "applied": 0,
                "dry_run": not getattr(args, "apply", False),
                "message": "no duplicate files found",
            }
        )
        return res

    keep_policy = getattr(args, "keep", "oldest") or "oldest"
    plan: list[dict[str, str]] = []
    for g in groups:
        ordered = list(g["files"])
        if keep_policy == "newest":
            ordered.sort(key=lambda p: p.stat().st_mtime, reverse=True)
        elif keep_policy == "shortest-path":
            ordered.sort(key=lambda p: (len(str(p)), str(p)))
        elif keep_policy == "longest-path":
            ordered.sort(key=lambda p: (-len(str(p)), str(p)))
        keep, dups = ordered[0], ordered[1:]
        for d in dups:
            plan.append({"action": "delete" if getattr(args, "apply", False) else "would_delete",
                         "keep": str(keep), "duplicate": str(d), "bytes": str(g["bytes_each"])})

    out = out_dir(getattr(args, "out", None), "dedupe")
    p_plan = unique_path(out / "dedupe_plan.csv")
    ensure_parent(p_plan)
    with open(p_plan, "w", encoding="utf-8-sig", newline="") as fh:
        w = csv.DictWriter(fh, fieldnames=["action", "duplicate", "keep", "bytes"])
        w.writeheader()
        w.writerows(plan)
    res.add_artifact(p_plan, "dedupe plan (review before applying)")

    applied = 0
    errors: list[str] = []
    if getattr(args, "apply", False):
        trash = out / "trash"
        mode = getattr(args, "mode", "trash") or "trash"
        for item in plan:
            src = Path(item["duplicate"])
            try:
                if mode == "delete":
                    src.unlink()
                else:
                    target = unique_path(trash / src.name)
                    ensure_parent(target)
                    shutil.move(str(src), str(target))
                    item["moved_to"] = str(target)
                applied += 1
            except OSError as exc:
                errors.append(f"{src}: {exc}")
        if errors:
            res.warn(f"{len(errors)} file(s) could not be processed")
    else:
        res.warn(
            f"dry run: {len(plan)} duplicate(s) would be removed, "
            f"reclaiming {human_bytes(sum(int(p['bytes']) for p in plan))}. "
            f"Re-run with --apply to act."
        )

    res.data.update(
        {
            "duplicate_groups": len(groups),
            "duplicate_files": len(plan),
            "reclaimable_bytes": sum(int(p["bytes"]) for p in plan),
            "reclaimable_human": human_bytes(sum(int(p["bytes"]) for p in plan)),
            "keep_policy": keep_policy,
            "applied": applied,
            "dry_run": not getattr(args, "apply", False),
            "mode": getattr(args, "mode", "trash"),
            "sample": plan[:20],
            "errors": errors[:20],
        }
    )
    return res


# --------------------------------------------------------------------------
# command: rename
# --------------------------------------------------------------------------
def _render_template(template: str, p: Path, index: int, *, prefix: str = "", suffix: str = "",
                     find: str | None = None, replace: str = "") -> str:
    st = p.stat()
    mtime = datetime.fromtimestamp(st.st_mtime)
    stem = p.stem
    if find:
        stem = re.sub(find, replace, stem)
    fields = {
        "name": stem,
        "ext": p.suffix.lstrip("."),
        "n": str(index),
        "num": str(index),
        "nn": f"{index:02d}",
        "nnn": f"{index:03d}",
        "date": mtime.strftime("%Y%m%d"),
        "time": mtime.strftime("%H%M%S"),
        "datetime": mtime.strftime("%Y%m%d_%H%M%S"),
        "year": mtime.strftime("%Y"),
        "month": mtime.strftime("%m"),
        "day": mtime.strftime("%d"),
        "parent": p.parent.name,
    }
    out = template
    for k, v in fields.items():
        out = out.replace("{" + k + "}", v)
    out = re.sub(r"\{[^}]*\}", "", out)  # drop unknown placeholders rather than emit braces
    return f"{prefix}{out}{suffix}"


def cmd_rename(args) -> Result:
    res = Result("rename")
    roots = resolve_roots(args.input)
    files: list[Path] = []
    for root in roots:
        if root.is_file():
            files.append(root)
        else:
            files.extend(walk_files(root, recursive=not getattr(args, "shallow", False),
                                    include=_split(getattr(args, "include", None)),
                                    exclude=_split(getattr(args, "exclude", None))))
    files.sort(key=lambda p: (str(p.parent), p.name))
    template = getattr(args, "template", None)
    find = getattr(args, "find", None)
    if not template and not find:
        raise OfficeKitError("provide --template (e.g. '{date}_{nnn}') or --find/--replace")

    plan: list[dict[str, Any]] = []
    used: set[str] = set()
    for idx, p in enumerate(files, start=1):
        try:
            st = p.stat()
        except OSError:
            continue
        if template:
            new_stem = _render_template(
                template, p, idx,
                prefix=getattr(args, "prefix", "") or "",
                suffix=getattr(args, "suffix", "") or "",
            )
        else:
            new_stem = _render_template("{name}", p, idx,
                                        find=find, replace=getattr(args, "replace", "") or "")
        if not new_stem.strip():
            new_stem = p.stem
        new_ext = p.suffix if getattr(args, "keep_extension", True) else ""
        candidate = p.with_name(f"{new_stem}{new_ext}")
        key = str(candidate.resolve()).lower()
        if key in used or (candidate.exists() and candidate != p):
            candidate = unique_path(candidate)
            key = str(candidate.resolve()).lower()
        used.add(key)
        if candidate == p:
            continue
        plan.append(
            {
                "source": str(p),
                "target": str(candidate),
                "old_name": p.name,
                "new_name": candidate.name,
                "bytes": st.st_size,
            }
        )

    out = out_dir(getattr(args, "out", None), "rename")
    p_plan = unique_path(out / "rename_plan.csv")
    ensure_parent(p_plan)
    with open(p_plan, "w", encoding="utf-8-sig", newline="") as fh:
        w = csv.DictWriter(fh, fieldnames=["old_name", "new_name", "source", "target", "bytes"])
        w.writeheader()
        w.writerows(plan)
    res.add_artifact(p_plan, "rename plan")

    applied, errors = 0, []
    if getattr(args, "apply", False):
        for item in plan:
            try:
                os.rename(item["source"], item["target"])
                applied += 1
            except OSError as exc:
                errors.append(f"{item['old_name']}: {exc}")
        if errors:
            res.warn(f"{len(errors)} file(s) failed to rename")
    else:
        res.warn(f"dry run: {len(plan)} file(s) would be renamed. Re-run with --apply to act.")

    res.data.update(
        {
            "scanned": len(files),
            "to_rename": len(plan),
            "applied": applied,
            "dry_run": not getattr(args, "apply", False),
            "template": template,
            "sample": plan[:25],
            "errors": errors[:20],
        }
    )
    return res


# --------------------------------------------------------------------------
# command: organize
# --------------------------------------------------------------------------
def cmd_organize(args) -> Result:
    res = Result("organize")
    src_root = Path(args.input).expanduser().resolve()
    if not src_root.is_dir():
        raise OfficeKitError(f"--input must be a directory for organize: {src_root}")
    dest_root = Path(getattr(args, "dest", None) or (src_root / "整理后_Organized")).resolve()
    if dest_root == src_root:
        raise OfficeKitError("destination must differ from the source directory")
    if dest_root.is_relative_to(src_root) and getattr(args, "apply", False):
        res.warn("destination is inside the source tree; it will be skipped during the scan")

    strategy = (getattr(args, "by", None) or "type").lower()
    if strategy not in ("type", "date", "year-month", "extension", "first-letter"):
        raise OfficeKitError("--by must be one of type, date, year-month, extension, first-letter")

    files = walk_files(src_root, recursive=not getattr(args, "shallow", False),
                       include=_split(getattr(args, "include", None)),
                       exclude=_split(getattr(args, "exclude", None)))
    files = [f for f in files if not str(f).startswith(str(dest_root))]

    plan: list[dict[str, Any]] = []
    for p in files:
        if p.name.lower() in ("desktop.ini", "thumbs.db"):
            continue
        cat = _category_dir(p, strategy)
        target = dest_root / cat / p.name
        resolution = "move" if getattr(args, "mode", "move") == "move" else "copy"
        if target.exists():
            resolution = "rename"
            target = unique_path(target)
        plan.append(
            {
                "source": str(p),
                "target": str(target),
                "category": cat,
                "bytes": p.stat().st_size,
                "resolution": resolution,
            }
        )

    out = out_dir(getattr(args, "out", None), "organize")
    p_plan = unique_path(out / "organize_plan.csv")
    ensure_parent(p_plan)
    with open(p_plan, "w", encoding="utf-8-sig", newline="") as fh:
        w = csv.DictWriter(fh, fieldnames=["category", "resolution", "source", "target", "bytes"])
        w.writeheader()
        w.writerows(plan)
    res.add_artifact(p_plan, "organize plan (review before applying)")

    applied, errors = 0, []
    if getattr(args, "apply", False):
        for item in plan:
            try:
                ensure_parent(item["target"])
                if getattr(args, "mode", "move") == "move":
                    shutil.move(item["source"], item["target"])
                else:
                    shutil.copy2(item["source"], item["target"])
                applied += 1
            except OSError as exc:
                errors.append(f"{item['source']}: {exc}")
        if errors:
            res.warn(f"{len(errors)} file(s) failed")
    else:
        res.warn(
            f"dry run: {len(plan)} file(s) would be organized into {dest_root} by {strategy}. "
            f"Re-run with --apply to act."
        )

    counts: dict[str, int] = defaultdict(int)
    sizes: dict[str, int] = defaultdict(int)
    for item in plan:
        counts[item["category"]] += 1
        sizes[item["category"]] += item["bytes"]

    res.data.update(
        {
            "source_root": str(src_root),
            "dest_root": str(dest_root),
            "strategy": strategy,
            "mode": getattr(args, "mode", "move"),
            "planned": len(plan),
            "applied": applied,
            "dry_run": not getattr(args, "apply", False),
            "buckets": {k: {"count": counts[k], "human": human_bytes(sizes[k])}
                        for k in sorted(counts, key=lambda x: -counts[x])},
            "sample": plan[:25],
            "errors": errors[:20],
        }
    )
    return res


def _category_dir(p: Path, strategy: str) -> str:
    if strategy == "type":
        return _suffix_category(p.suffix)
    if strategy == "extension":
        return (p.suffix.lstrip(".") or "无扩展名").lower()
    if strategy == "first-letter":
        ch = p.stem[:1].upper()
        return ch if ch.isalnum() else "#"
    mtime = datetime.fromtimestamp(p.stat().st_mtime)
    if strategy == "date":
        return mtime.strftime("%Y-%m-%d")
    return mtime.strftime("%Y-%m")


# --------------------------------------------------------------------------
# command: archive
# --------------------------------------------------------------------------
def cmd_archive(args) -> Result:
    import zipfile

    res = Result("archive")
    roots = resolve_roots(args.input)
    files: list[Path] = []
    base: Path | None = None
    for root in roots:
        if root.is_dir():
            base = base or root
            files.extend(walk_files(root, recursive=True,
                                    include=_split(getattr(args, "include", None)),
                                    exclude=_split(getattr(args, "exclude", None))))
        else:
            base = base or root.parent
            files.append(root)
    if not files:
        raise OfficeKitError("no files to archive")

    out = out_dir(getattr(args, "out", None), "archive")
    name = safe_stem(getattr(args, "name", None) or (base.name if base else "archive"))
    target = unique_path(out / f"{name}.zip")
    compression = zipfile.ZIP_DEFLATED if getattr(args, "compress", True) else zipfile.ZIP_STORED
    total = 0
    ensure_parent(target)
    with zipfile.ZipFile(target, "w", compression=compression, compresslevel=6) as z:
        for p in files:
            try:
                arc = str(p.relative_to(base)) if base and str(p).startswith(str(base)) else p.name
            except ValueError:
                arc = p.name
            z.write(p, arcname=arc)
            total += p.stat().st_size
    res.add_artifact(target, f"archive of {len(files)} files")
    res.data.update(
        {
            "files": len(files),
            "source_bytes": total,
            "source_human": human_bytes(total),
            "archive_bytes": target.stat().st_size,
            "archive_human": human_bytes(target.stat().st_size),
            "ratio": round(target.stat().st_size / total, 4) if total else None,
        }
    )
    return res
