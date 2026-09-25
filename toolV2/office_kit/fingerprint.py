"""工具指纹：**这份产物是哪个版本的工具做出来的？**

为什么必须有这个（[24 §4.4 C](../project_bank_adjust/24_歧义与模糊清单.md)）
------------------------------------------------------------------------
"工具不够用时怎么办"要走一条分诊：**共性问题 = 工具升级；特殊问题 = 复制到本会话工作区、
标注特殊版、不改 `D:\\DSH\\tool`**。

可是**没有指纹，这条分诊就落不了地**：

* 特殊版 fork 出去以后，**认不出它亲缘是谁**（改了哪几行？从哪个版本改的？）；
* 产物进了档案，**答不出"这是哪个版本的工具做的"**；
* 出了问题要复现，**不知道该装回哪一版**。

所以指纹 = **把"工具版本"变成可写进产物记录里的一串字符**。

怎么算
------
把 `office.py` 与 `office_kit/**.py` 的 **相对路径 + 各自 sha256** 排序后拼起来，再取一次摘要。
**只看源码**，不看 `__pycache__`、不看数据、不看产物——
这样"同样的源码 = 同样的指纹"，与时间、机器、运行次数无关。

⚠️ **诚实边界**：指纹**只能发现"源码变了"**，**不能证明源码是清白的**
（谁都能改完源码再算一次指纹）。它解决的是"**认亲**"与"**复现**"，不是"**防篡改**"。
"""
from __future__ import annotations

import hashlib
from pathlib import Path

#: 参与指纹的文件：入口 + 包里的 Python 源码
ROOT = Path(__file__).resolve().parent.parent
PACKAGE = ROOT / "office_kit"
ENTRY = ROOT / "office.py"


def _digest(text: str) -> str:
    try:
        return hashlib.new("sm3", text.encode("utf-8")).hexdigest()
    except Exception:  # noqa: BLE001 - 算法可配置（D-5），退到 sha256
        return hashlib.sha256(text.encode("utf-8")).hexdigest()


def _source_files() -> list[Path]:
    files = [ENTRY] if ENTRY.exists() else []
    if PACKAGE.is_dir():
        files += sorted(p for p in PACKAGE.glob("*.py") if p.name != "__init__.py" or True)
    return sorted(p for p in files if p.is_file())


def tool_fingerprint(*, short: bool = False) -> dict:
    """算一次工具指纹。

    返回 `{"digest": "…", "short": "…", "files": N, "algo": "sm3|sha256"}`
    或 `{"digest": None, "error": "…"}`（算不出来时**如实报错，不编一个**）。
    """
    try:
        parts: list[str] = []
        for p in _source_files():
            h = hashlib.sha256(p.read_bytes()).hexdigest()
            rel = str(p.relative_to(ROOT)).replace("\\", "/")
            parts.append("%s %s" % (rel, h))
        if not parts:
            return {"digest": None, "short": None, "files": 0,
                    "error": "找不到任何源码文件（%s）" % ROOT}
        material = "\n".join(parts)
        full = _digest(material)
        algo = "sm3" if len(full) == 64 else "sha256"
        try:
            hashlib.new("sm3")
        except Exception:  # noqa: BLE001
            algo = "sha256"
        return {"digest": full, "short": full[:12], "files": len(parts), "algo": algo,
                "root": str(ROOT)}
    except Exception as exc:  # noqa: BLE001
        return {"digest": None, "short": None, "files": 0,
                "error": "%s: %s" % (type(exc).__name__, exc)}


def short_fingerprint() -> str:
    """一行短指纹，写进 `fill_op.preflight_json` 与事件载荷里用。"""
    fp = tool_fingerprint()
    return fp.get("short") or "unknown"


def file_hashes() -> dict[str, str]:
    """逐文件 sha256——fork 时写进 `FORK.md`，好逐行比对"改了哪几个文件"。"""
    out: dict[str, str] = {}
    for p in _source_files():
        out[str(p.relative_to(ROOT)).replace("\\", "/")] = \
            hashlib.sha256(p.read_bytes()).hexdigest()
    return out
