r"""工作区（工作根）——**每次由使用者指定，随便哪个文件夹**。

口径（主人第 18 轮已定，唯一权威：[26 §2.9]）
-----------------------------------------------
**工作区不是固定位置，是"这一次活儿干在哪个文件夹里"。**你指一个文件夹，
库、源文件副本、产物、日志、临时区**全都长在那个文件夹里**，一眼看得见、拷走就是全部。

```
<工作区>\
  ├─ .office-workroot.json     标记文件：声明"这里是一个工作区"（命令靠它认领）
  ├─ in\<批次号>\              源文件的整份副本（只读）
  ├─ out\<批次号>\             产物
  │    ├─ <主体名>\             填好的目标文件（按主体分，26 §2.9）
  │    └─ _报告\<命令>\         人看的报告
  ├─ db\workflow.db           数据库（SCHEMA v2）
  ├─ logs\<YYYYMMDD>.log      运行日志（按天滚动）
  └─ work\<批次号>\            临时区（可清理）
```

**怎么认领一个工作区**（`resolve` 的顺序，**不偷偷落到当前目录**）：

1. 命令上显式给了 `--db <文件>` → 就用它（单文件调试用，最优先）
2. 命令上显式给了 `--work <文件夹>` → 用它
3. 环境变量 `DSH_OFFICE_WORK` → 用它
4. 从当前目录**逐级往上**找 `.office-workroot.json` → 找到就用那个
5. 都没有 → 🔴 **报错，并告诉你两种给法**（`work --init --at <文件夹>` / `--work <文件夹>`）

⚠️ **为什么第 5 条是报错而不是"就用当前目录"**：银行场景里"文件不知道跑哪去了"
是最糟的一类事故。宁可不干，也不能把库悄悄建在一个谁也没指定的地方。
"""
from __future__ import annotations

import json
import os
import re
import shutil
import time
from dataclasses import dataclass
from pathlib import Path

MARKER_NAME = ".office-workroot.json"
MARKER_VERSION = 1
ENV_VAR = "DSH_OFFICE_WORK"

#: 工作区里的五个子目录（[26 §2.9]）
SUBDIRS = ("in", "out", "db", "logs", "work")
#: 报告统一放在 `out\<批次号>\_报告\<命令>\`，跟产物分开
REPORT_DIR = "_报告"

BATCH_RE = re.compile(r"^\d{8}-\d{2}$")


class WorkrootError(RuntimeError):
    """工作区没指定、找不到、或者不是工作区。**消息必须能直接照做。**"""


@dataclass
class Workroot:
    """一个已认领的工作区。`how` 记下"是怎么认出来的"，好在报告里如实交代。"""

    path: Path
    how: str

    # ---- 固定位置 ------------------------------------------------------
    @property
    def marker(self) -> Path:
        return self.path / MARKER_NAME

    @property
    def db(self) -> Path:
        return self.path / "db" / "workflow.db"

    @property
    def inbox(self) -> Path:
        r"""源文件副本的根（`in\<批次号>\`）。"""
        return self.path / "in"

    @property
    def outbox(self) -> Path:
        r"""产物根（`out\<批次号>\`）。"""
        return self.path / "out"

    @property
    def logs(self) -> Path:
        return self.path / "logs"

    @property
    def work(self) -> Path:
        return self.path / "work"

    # ---- 按批次取 ------------------------------------------------------
    def in_dir(self, batch_no: str) -> Path:
        return self.inbox / batch_no

    def out_dir(self, batch_no: str) -> Path:
        return self.outbox / batch_no

    def subject_dir(self, batch_no: str, subject: str) -> Path:
        r"""产物按**主体**分目录（`out\<批次号>\<主体名>\`，[26 §2.9]）。"""
        return self.out_dir(batch_no) / safe_name(subject)

    def report_dir(self, batch_no: str, command: str) -> Path:
        return self.out_dir(batch_no) / REPORT_DIR / command

    def work_dir(self, batch_no: str) -> Path:
        return self.work / batch_no

    def log_file(self, day: str | None = None) -> Path:
        return self.logs / ("%s.log" % (day or time.strftime("%Y%m%d")))

    # ---- 动作 ----------------------------------------------------------
    def ensure(self) -> "Workroot":
        for sub in SUBDIRS:
            (self.path / sub).mkdir(parents=True, exist_ok=True)
        return self

    def log(self, text: str, *, day: str | None = None) -> None:
        r"""往 `logs\<YYYYMMDD>.log` 追加一行。**日志写不进去也不许让活儿失败。**

        ⚠️ **只有真正的工作区才记日志**（有标记文件）。用 `--db` 指了一个不在任何工作区里的库
        （单文件调试）时**什么都不建**——否则会在人家目录里凭空掉一个 `logs\` 文件夹，
        那正是"东西四处散落"的来源（实测：`examples\mvp-demo\logs\` 就是这么冒出来的）。
        """
        try:
            if not self.marker.exists():
                return
            self.logs.mkdir(parents=True, exist_ok=True)
            stamp = time.strftime("%Y-%m-%d %H:%M:%S")
            with open(self.log_file(day), "a", encoding="utf-8") as fh:
                fh.write("%s  %s\n" % (stamp, text))
        except Exception:  # noqa: BLE001
            pass

    def describe(self) -> dict:
        return {"工作区": str(self.path), "怎么认出来的": self.how,
                "数据库": str(self.db), "库存在": self.db.exists(),
                "标记文件": str(self.marker), "标记存在": self.marker.exists()}


def safe_name(name: str) -> str:
    """主体名做目录名：去掉路径分隔符等不能做文件名的字符。"""
    s = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", str(name or "")).strip(" .")
    return s[:60] or "（未归属）"


# --------------------------------------------------------------------------
# 建 / 找
# --------------------------------------------------------------------------
def init_workroot(path: str | Path, *, note: str | None = None,
                  allow_existing: bool = True) -> dict:
    """把一个文件夹变成工作区：建五个子目录 + 写标记文件。**不删任何东西。**"""
    p = Path(path).expanduser()
    if p.exists() and not p.is_dir():
        raise WorkrootError("这个路径是个文件，不是文件夹：%s" % p)
    marker = p / MARKER_NAME
    if marker.exists() and not allow_existing:
        raise WorkrootError("这已经是一个工作区了：%s" % p)
    p.mkdir(parents=True, exist_ok=True)
    made: list[str] = []
    for sub in SUBDIRS:
        d = p / sub
        if not d.exists():
            d.mkdir(parents=True, exist_ok=True)
            made.append(sub)
    created = not marker.exists()
    if created or allow_existing:
        payload = {"marker_version": MARKER_VERSION, "kind": "office-workroot",
                   "created_at": _iso(), "note": note}
        if marker.exists():
            old = _load_marker(marker)
            payload["created_at"] = old.get("created_at") or payload["created_at"]
            payload["note"] = note if note is not None else old.get("note")
        marker.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    wr = Workroot(p.resolve(), "work --init")
    wr.ensure()
    return {"workroot": str(wr.path), "marker": str(marker), "created_marker": created,
            "created_dirs": made, "db": str(wr.db), "db_exists": wr.db.exists(),
            "dirs": {s: str(wr.path / s) for s in SUBDIRS}}


def _load_marker(path: Path) -> dict:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception:  # noqa: BLE001
        return {}


def _iso() -> str:
    from datetime import datetime, timezone

    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def find_workroot(start: str | Path | None = None,
                  *, max_up: int = 24) -> tuple[Path, str] | None:
    """从 `start`（默认当前目录）逐级往上找标记文件。"""
    cur = Path(start or Path.cwd()).expanduser()
    try:
        cur = cur.resolve()
    except Exception:  # noqa: BLE001
        return None
    if cur.is_file():
        cur = cur.parent
    for _ in range(max_up):
        if (cur / MARKER_NAME).exists():
            return cur, "从当前目录往上找到标记文件 %s" % MARKER_NAME
        if cur.parent == cur:
            break
        cur = cur.parent
    return None


def is_workroot(path: str | Path) -> bool:
    p = Path(path).expanduser()
    return (p / MARKER_NAME).exists()


def resolve(*, db: str | Path | None = None, work: str | Path | None = None,
            start: str | Path | None = None, must_exist: bool = False) -> tuple[Workroot, Path]:
    """按文件的五步顺序认领工作区，返回 `(工作区, 库文件)`。

    ⚠️ 第 5 步是**报错**，不是"就用当前目录"——见文件头的理由。
    """
    if db:
        p = Path(db).expanduser()
        if must_exist and not p.exists():
            raise WorkrootError("数据库不存在：%s" % p)
        # 库文件在某个工作区里 → 连带把工作区也认出来（产物好落对地方）
        found = find_workroot(p.parent)
        if found:
            return Workroot(found[0], "由 --db 反查（%s）" % found[1]), p
        return Workroot(p.parent.resolve(), "由 --db 指定，不在任何工作区内"), p

    if work:
        p = Path(work).expanduser().resolve()
        if not p.is_dir():
            raise WorkrootError(
                "这个工作区文件夹不存在：%s\n"
                "先建一个：  office.py work --init --at \"%s\"" % (p, p))
        if not is_workroot(p):
            raise WorkrootError(
                "这个文件夹还不是工作区（没有 %s）：%s\n"
                "把它变成工作区：  office.py work --init --at \"%s\""
                % (MARKER_NAME, p, p))
        wr = Workroot(p, "由 --work 指定")
        return wr, wr.db

    env = os.environ.get(ENV_VAR)
    if env:
        p = Path(env).expanduser().resolve()
        if not is_workroot(p):
            raise WorkrootError("环境变量 %s 指向的文件夹不是工作区：%s" % (ENV_VAR, p))
        wr = Workroot(p, "由环境变量 %s 指定" % ENV_VAR)
        return wr, wr.db

    found = find_workroot(start)
    if found:
        wr = Workroot(found[0], found[1])
        return wr, wr.db

    raise WorkrootError(
        "🔴 没有指定工作区，也没有在上级目录里找到 %s。\n"
        "**不会自动在当前目录建库**（免得文件不知道跑哪去了）。两种给法：\n"
        "  ① 就把当前目录当工作区（最常用）： office.py work --init\n"
        "     当前目录是： %s\n"
        "  ② 指到别处： office.py work --init --at \"D:\\某个文件夹\"\n"
        "     或给已有工作区： 在命令后加 --work \"D:\\某个文件夹\"\n"
        "（也可以直接 --db <库文件> 只指定库）" % (MARKER_NAME, Path.cwd()))


def session_report(wr: "Workroot | None" = None) -> dict:
    """把"工作区"与"这次 DSH 会话"的关系摊开，**让使用者一眼看出对不对得上**。

    ⚠️ **诚实边界**：DSH 目前**只往子进程传 4 个环境变量**
    （`DSH_HOME` / `DSH_SESSION_ID` / `DSH_SHELL` / `DSH_WEB_URL`），
    **不传会话工作区、也不传权限档位**。所以本工具**读不到**"DSH 认为你的工作区是哪、
    权限开到哪"。这里能确定的只有：
      * 当前目录在哪、它在不在工作区里面（这两件是本进程自己算的）；
      * `DSH_HOME` 是多少（环境变量）。
    **权限档位必须由代理在对话里说清楚**——见 SKILL.md「动手前先对齐三件事」。
    """
    cwd = Path.cwd()
    inside = None
    if wr is not None:
        try:
            cwd.resolve().relative_to(wr.path.resolve())
            inside = True
        except Exception:  # noqa: BLE001
            inside = False
    return {
        "当前目录": str(cwd),
        "当前目录在工作区内": inside,
        "工作区": str(wr.path) if wr else None,
        "DSH_HOME": os.environ.get("DSH_HOME"),
        "DSH_SESSION_ID": os.environ.get("DSH_SESSION_ID"),
        "能读到的 DSH 环境变量": sorted(k for k in os.environ if k.startswith("DSH_")),
        "权限档位": "**读不到**——DSH 不把它传给子进程，必须由代理在对话里说明",
    }



def ensure_layout(wr: Workroot) -> Workroot:
    return wr.ensure()


def batch_dir_name(batch_no: str | None) -> str:
    """产物的批次子目录名。没有批次时用一个明说的名字，不编造批次号。"""
    if batch_no and BATCH_RE.match(batch_no):
        return batch_no
    return "未编号批次"


# --------------------------------------------------------------------------
# 越界告警（**中文，必须说人话**）
# --------------------------------------------------------------------------
def outside_warning(args, path: str | Path) -> str | None:
    """这个落点**超出工作区**了吗？超出就返回一句中文提示（没有就返回 None）。

    为什么只告警不拦（[24 §4.4 ③](../project_bank_adjust/24_歧义与模糊清单.md)）：

    * **文件夹不是边界**——有文件系统权限的进程能往任何地方写，工具拦不住；
    * 真正拦得住的是 **DSH 的权限档位**（"仅可查看 / 工作区内修改 / 完全权限"）；
    * 工具能做、也应该做的是：**发现越界就说出来**，别让它悄悄发生。

    ⚠️ 只在**已经认领到真正工作区**（有标记文件）时才判——用 `--db` 单文件调试时
    没有工作区可言，不该到处报"越界"。
    """
    wr = getattr(args, "_workroot", None)
    if wr is None or not is_workroot(wr.path):
        return None
    try:
        Path(path).expanduser().resolve().relative_to(wr.path.resolve())
        return None
    except Exception:  # noqa: BLE001
        return ("⚠️ **越界提示**：`%s` **不在工作区 `%s` 里面**。\n"
                "    本次会写到工作区之外——请确认这是你要的：\n"
                "    · 如果只是想换个地方放产物，建议改到工作区内（`out\\<批次号>\\…`）；\n"
                "    · 如果确实要写出去，**权限档位得够**（工具拦不住，只能提醒你）；\n"
                "    · 另外注意：**别的会话的工作区不要碰**（一个会话一个工作区）。"
                % (path, wr.path))


def outside_warnings(args, *paths) -> list[str]:
    """把越界告警**按目录合并**：同一个目录只报一次，别刷屏。

    （实测：一次 `db-fill` 的 `--out` 会同时带出 `fill_plan.md` 与 `fill_plan.json` 两条，
    内容一模一样只是文件名不同——报两遍等于噪音，反而让人不去看。）
    """
    wr = getattr(args, "_workroot", None)
    if wr is None or not is_workroot(wr.path):
        return []
    outside: dict[str, list[str]] = {}
    for p in paths:
        if p is None:
            continue
        try:
            Path(p).expanduser().resolve().relative_to(wr.path.resolve())
            continue  # 在工作区里，不用管
        except Exception:  # noqa: BLE001
            pass
        pp = Path(p).expanduser()
        outside.setdefault(str(pp.parent), []).append(pp.name)
    if not outside:
        return []
    lines = []
    for d, names in outside.items():
        shown = "、".join("`%s`" % n for n in names[:2])
        more = "（同目录另外 %d 个文件）" % (len(names) - 2) if len(names) > 2 else ""
        lines.append(
            "⚠️ **越界提示**：`%s` 下的 %s%s **不在工作区 `%s` 里面**。\n"
            "    本次会写到工作区之外——请确认这是你要的：\n"
            "    · 只是想换个地方放产物 → 建议改到工作区内（`out\\<批次号>\\…`）；\n"
            "    · 确实要写出去 → **权限档位得够**（工具拦不住，只能提醒你）；\n"
            "    · 另外注意：**别的会话的工作区不要碰**（一个会话一个工作区）。"
            % (d, shown, more, wr.path))
    return lines


# --------------------------------------------------------------------------
# 建议一个工作区（**给用户审核用**）
# --------------------------------------------------------------------------
def _name_from(text: str | None) -> str | None:
    if not text:
        return None
    stem = re.sub(r"[\\/]+$", "", str(text)).split("\\")[-1].split("/")[-1]
    stem = re.sub(r"\.[A-Za-z0-9]{1,5}$", "", stem)
    stem = re.sub(r"^[\d.、\-\s]+", "", stem).strip()
    stem = re.sub(r"[<>:\"/\\|?*\x00-\x1f]", "", stem)
    stem = re.sub(r"\s+", "", stem)
    return stem[:24] or None


def suggest(*, cwd: str | Path | None = None, for_what: str | None = None,
            day: str | None = None) -> dict:
    """给使用者**一个具体建议**（目录 + 文件夹名 + 为什么这么建议），让他一句话就能拍板。

    规则（**保守**）：

    * **父目录 = 当前会话的工作目录**——这样 DSH 的权限档位「工作区内修改」就够用，**不用提权**；
    * ⚠️ **绝不建议放在工具包里面**（`D:\\DSH\\tool\\office`）——那是**程序目录**，
      往里放工作区就是**代码和数据混在一起**：备份、升级、比对指纹全乱。
      当前目录要是在工具包里，就往上升到工具的**上一级**再建议，并把原因说出来；
    * **名字 = `<事由>-<YYYYMMDD>`**，事由从材料名推（推不出就用 `办公工作区`）；
    * 名字**已存在就加 `-2`、`-3`**，不覆盖、不合并；
    * 同时给出"如果你想放别处"的一句话写法。
    """
    toolkit = Path(__file__).resolve().parent.parent
    base = Path(cwd or Path.cwd()).expanduser()
    try:
        base = base.resolve()
    except Exception:  # noqa: BLE001
        pass

    note = None
    inside_toolkit = False
    try:
        base.relative_to(toolkit)
        inside_toolkit = True
    except Exception:  # noqa: BLE001
        inside_toolkit = False
    if inside_toolkit:
        note = ("当前目录 `%s` **在工具包里面**（那是程序目录），"
                "放工作区会把**代码和数据混在一起**——所以我建议放到 `%s` 下面。"
                % (base, toolkit.parent.parent))
        base = toolkit.parent.parent

    name = _name_from(for_what) or "办公工作区"
    stamp = day or time.strftime("%Y%m%d")
    candidate = "%s-%s" % (name, stamp)
    full = base / candidate
    n = 1
    while full.exists():
        n += 1
        candidate = "%s-%s-%d" % (name, stamp, n)
        full = base / candidate
    info = {
        "建议父目录": str(base),
        "建议文件夹名": candidate,
        "建议完整路径": str(full),
        "为什么放这里": "放在**当前会话的工作目录**下面 → DSH 权限档位选「工作区内修改」就够，"
                        "**不用提权到完全权限**",
        "如果你想放别处": "office.py work --init --at \"D:\\你想要的路径\\%s\"" % candidate,
        "要不要现在建": "office.py work --init --at \"%s\"" % full,
    }
    if note:
        info["注意"] = note
    return info


def render_suggest(info: dict) -> str:
    L = [
        "# 建议的工作区（**请你审核**）",
        "",
        "| 项 | 建议 |",
        "| --- | --- |",
        "| 放在哪 | `%s` |" % info["建议父目录"],
        "| 文件夹叫什么 | **`%s`** |" % info["建议文件夹名"],
        "| 完整路径 | `%s` |" % info["建议完整路径"],
        "",
        "**为什么这么建议**：%s" % info["为什么放这里"],
        "",
        "你回一句「可以」，我就执行：",
        "",
        "```",
        info["要不要现在建"],
        "```",
        "",
        "想放别处也行，把路径告诉我（或者直接用下面这条）：",
        "",
        "```",
        info["如果你想放别处"],
        "```",
        "",
        "> 📌 建好以后，这个文件夹里会有 `in\\ out\\ db\\ logs\\ work\\` 五个子目录和"
        "一个 `.office-workroot.json` 标记文件；**库、源文件副本、产物、报告、日志全在它里面**，"
        "拷走就是全部。",
        "",
        "> ⚠️ **在你确认之前，工具不会产出任何文件**——所以不会有「先落在别处、回头再搬」的麻烦"
        "（第 21 批已定：**不设「临时文件夹」这个概念**）。",
    ]
    if info.get("注意"):
        L += ["", "> ⚠️ **注意**：%s" % info["注意"]]
    return "\n".join(L) + "\n"


def brief() -> str:
    """代理**动手前该对使用者说的那段话**，原样打印，免得每次现编、编漏。"""
    wr = None
    try:
        found = find_workroot()
        if found:
            wr = Workroot(found[0], found[1])
    except Exception:  # noqa: BLE001
        wr = None
    fp = _fp()
    L = ["## 动手前对齐（请把下面三行原样告诉使用者）", ""]
    L.append("1. **工具**：我用本机的 `office_kit` 工具包（`%s`）来做，不临时写脚本。"
             "工具指纹 `%s`。" % (ROOT_OF_TOOLKIT, fp))
    if wr is not None:
        L.append("2. **工作区**：**这次的活儿干在 `%s`**。库、源文件副本、产物、报告、日志都会落在它下面。"
                 % wr.path)
    else:
        L.append("2. **工作区**：**还没定**——我会先给你一个建议路径，你点头之后再开工。"
                 "（在你确认之前，工具**不会产出任何文件**）")
    L.append("3. **权限**：这一步**只需要**：读工具包 + 在工作区内读写。"
             "**不需要**动工作区之外的东西。")
    L.append("")
    L.append("> ⚠️ 本工具**读不到** DSH 的权限档位（DSH 不往子进程传），所以这句必须由你说。")
    return "\n".join(L)


ROOT_OF_TOOLKIT = str(Path(__file__).resolve().parent.parent)


def _fp() -> str:
    try:
        from .fingerprint import short_fingerprint

        return short_fingerprint()
    except Exception:  # noqa: BLE001
        return "unknown"


# --------------------------------------------------------------------------
# 分诊（F）：共性问题 → 工具升级；特殊问题 → fork 到本会话工作区
# --------------------------------------------------------------------------
TRIAGE_KINDS = ("common", "special")
TRIAGE_LABEL = {"common": "共性问题（工具升级）", "special": "特殊问题（本会话特殊版）"}


def triage(args) -> dict:
    """把"这是共性问题还是特殊问题"这个判断**记下来**，并按结论执行。

    * `common`  → 工具升级。**不自动改代码**，只把结论与理由留痕，
      并提示"改完要跑六套自测"。改代码本身要在**能改代码的会话**里做。
    * `special` → 把工具**复制**到 `<工作区>\\工具\\<名字>-特殊版-YYYYMMDD\\`，
      写 `FORK.md`，**绝不改 `D:\\DSH\\tool`**。

    两个事件都用 **[06 §2.1]** 已有的 `hook/invoked`（收到一次动作请求）与
    `hook/result`（判定返回）——**不造新事件名**。
    """
    from .fingerprint import file_hashes, tool_fingerprint

    kind = getattr(args, "kind", None)
    if kind not in TRIAGE_KINDS:
        raise WorkrootError("必须给 --kind common 或 --kind special（共性问题 / 特殊问题）")
    reason = (getattr(args, "reason", None) or "").strip()
    if not reason:
        raise WorkrootError("必须写明理由（--reason）：为什么判成共性问题 / 特殊问题")
    by = (getattr(args, "by", None) or "user").strip()
    fp = tool_fingerprint()

    wr = None
    try:
        wr, db = resolve(work=getattr(args, "work", None))
    except WorkrootError:
        wr, db = None, None

    info: dict = {"kind": kind, "决策": TRIAGE_LABEL[kind], "理由": reason, "批准人": by,
                  "工具指纹": fp.get("short"), "工作区": str(wr.path) if wr else None}
    events_written: list[str] = []

    if kind == "special":
        if wr is None or not is_workroot(wr.path):
            raise WorkrootError(
                "特殊问题**必须**复制到本会话的工作区里，可现在没有工作区。\n"
                "先定工作区： office.py work --init --at \"<文件夹>\"")
        fork_name = (getattr(args, "fork_name", None) or "特殊版").strip()
        day = time.strftime("%Y%m%d")
        dest = wr.path / "工具" / ("%s-%s" % (safe_name(fork_name), day))
        n = 1
        while dest.exists():
            n += 1
            dest = wr.path / "工具" / ("%s-%s-%d" % (safe_name(fork_name), day, n))
        dest.parent.mkdir(parents=True, exist_ok=True)
        shutil.copytree(Path(__file__).resolve().parent.parent, dest,
                        ignore=shutil.ignore_patterns("__pycache__", "*.pyc", "_office_out"))
        lines = [
            "# 特殊处理版本（**不是官方工具，不要当成官方版用**）", "",
            "| 项 | 值 |", "| --- | --- |",
            "| 为什么 fork（判为特殊问题） | %s |" % reason,
            "| 谁批的 | %s |" % by,
            "| fork 时间 | %s |" % time.strftime("%Y-%m-%d %H:%M:%S"),
            "| **原始工具指纹** | `%s`（%s，%d 个文件）"
            % (fp.get("digest"), fp.get("algo"), fp.get("files")),
            "| 原始工具位置 | `%s` |" % ROOT_OF_TOOLKIT,
            "| 本副本位置 | `%s` |" % dest,
            "| 本会话工作区 | `%s` |" % wr.path,
            "",
            "## 改了哪几行（**fork 完请如实填**）", "",
            "> 改完这一节要填上：改了哪个文件、哪一段、为什么。",
            "> 并且**重新算一次本副本的指纹**填在最下面。", "",
            "- （待填）", "",
            '## 原始版本逐文件 sha256（用来比对「我到底改了什么」）', "",
            "| 文件 | sha256 |", "| --- | --- |",
        ]
        for rel, h in sorted(file_hashes().items()):
            lines.append("| `%s` | `%s` |" % (rel, h))
        lines += ["", "## 本副本指纹（改完请重算并填这里）", "",
                  "```", "python office.py capabilities", "```", ""]
        (dest / "FORK.md").write_text("\n".join(lines) + "\n", encoding="utf-8")
        info.update({"fork_path": str(dest), "fork_md": str(dest / "FORK.md"),
                     "说明": "已复制到本会话工作区并写了 FORK.md；"
                             "**D:\\DSH\\tool 一个字都没改**。"
                             "以后这份活儿要跑**副本里的** office.py。"})
    else:
        info["说明"] = ("判为共性问题 → 这是**工具升级**。**本命令不改代码**：\n"
                        "① 改代码要在**能改代码的那个会话**里做（工作区设成 `%s`）；\n"
                        "② 改完**必须跑六套自测**（缺一不可），结果随交付；\n"
                        "③ 改完工具指纹会变，新产物会带新指纹。" % ROOT_OF_TOOLKIT)

    # ---- 留痕：hook/invoked + hook/result（06 §2.1 已有的枚举） ----
    if wr is not None and is_workroot(wr.path):
        try:
            from .store_v2 import StoreV2, allocate_batch

            store = StoreV2(wr.db, actor=by)
            bno = store.current_batch() or allocate_batch(store.conn)
            store.event("hook/invoked", batch_no=bno, target="临时编程",
                        payload={"请求": "工具不满足需求，需要临时处理",
                                 "判为": TRIAGE_LABEL[kind], "理由": reason,
                                 "工具指纹": fp.get("short"), "来源": "会话"},
                        actor=by)
            store.event("hook/result", batch_no=bno, target="临时编程",
                        payload={"走向": ("暂停问使用者 → 已批准 fork 到本会话工作区"
                                          if kind == "special" else
                                          "暂停问使用者 → 已批准按共性问题做工具升级"),
                                 "kind": kind, "批准人": by,
                                 "fork_path": info.get("fork_path"),
                                 "工具指纹": fp.get("short")},
                        actor=by)
            events_written = ["hook/invoked", "hook/result"]
            store.close()
        except Exception as exc:  # noqa: BLE001
            info["留痕失败"] = "%s: %s" % (type(exc).__name__, exc)
    else:
        info["留痕失败"] = "没有工作区，事件链无处可写（结论只落在 FORK.md / 本报告里）"
    info["事件"] = events_written
    return info


def render_triage(info: dict) -> str:
    L = ["# 临时编程分诊结果", "", "| 项 | 值 |", "| --- | --- |"]
    for k, v in info.items():
        if k == "说明":
            continue
        L.append("| %s | %s |" % (k, v))
    L += ["", "## 结论", "", str(info.get("说明") or ""), ""]
    if info.get("kind") == "special":
        L += ["## 下一步（照做，别跳步）", "",
              "1. **改副本，不改原版**：以后这份活儿跑 `%s` 里的 `office.py`；"
              % info.get("fork_path"),
              "2. 改完把 `FORK.md` 的「改了哪几行」填上，并重算本副本指纹；",
              "3. 交付时说明：**这份产物是特殊版做的，原始工具指纹是 `%s`**。"
              % info.get("工具指纹"), ""]
    else:
        L += ["## 下一步（照做，别跳步）", "",
              "1. 到**能改代码的会话**里改 `%s`；" % ROOT_OF_TOOLKIT,
              "2. 改完跑六套自测：`selftest.py` / `selftest_v2.py` / `selftest_entity_key.py` /"
              " `selftest_merge.py` / `selftest_pipeline_v2.py` / `selftest_workroot.py`；",
              "3. 自测结果随交付一起给使用者。", ""]
    return "\n".join(L) + "\n"




# --------------------------------------------------------------------------
# 命令层公共入口
# --------------------------------------------------------------------------
def open_for_command(args) -> Workroot:
    """给一条 `db-*` 命令认领工作区。**认不出来就报错，不猜。**"""
    wr, db = resolve(db=getattr(args, "db", None), work=getattr(args, "work", None))
    # 只有**真正的工作区**（有标记文件）才铺那五个目录。
    # 用 `--db` 指了一个不在工作区里的库（单文件调试）时，不去人家目录里乱建东西。
    if is_workroot(wr.path):
        wr.ensure()
    # 把认领结果放回 args，后面的代码统一读 args._db / args._workroot
    args._workroot = wr
    args._db = db
    return wr


def db_path(args) -> str:
    """命令该用哪个库。`--db` 优先；否则 `--work`；否则往上找标记文件。"""
    if getattr(args, "_db", None):
        return str(args._db)
    open_for_command(args)
    return str(args._db)


def report_dir_for(args, command: str, batch_no: str | None = None) -> Path:
    """报告落哪：显式 `--out` 优先；否则 `<工作区>\\out\\<批次>\\_报告\\<命令>\\`。"""
    from .common import out_dir

    explicit = getattr(args, "out", None)
    if explicit:
        return out_dir(explicit, command)
    wr = getattr(args, "_workroot", None)
    if wr is None:
        try:
            wr = open_for_command(args)
        except WorkrootError:
            return out_dir(None, command)
    d = wr.report_dir(batch_dir_name(batch_no), command)
    d.mkdir(parents=True, exist_ok=True)
    return d


def artifact_dir_for(args, batch_no: str, subject: str) -> Path:
    r"""填好的目标文件落哪。

    * **显式给了 `--out`** → `<out>\<批次>\<主体名>\`（使用者的话最大）；
    * 在工作区里干活 → `<工作区>\out\<批次>\<主体名>\`（[26 §2.9]）；
    * 都没有 → 当前目录下的 `产物\<批次>\<主体名>\`。
    """
    explicit = getattr(args, "out", None)
    if explicit:
        d = Path(explicit).expanduser() / batch_dir_name(batch_no) / safe_name(subject)
        d.mkdir(parents=True, exist_ok=True)
        return d
    wr = getattr(args, "_workroot", None)
    if wr is None or not is_workroot(wr.path):
        try:
            wr = open_for_command(args)
        except WorkrootError:
            wr = None
    if wr is None or not is_workroot(wr.path):
        d = Path.cwd() / "产物" / batch_dir_name(batch_no) / safe_name(subject)
        d.mkdir(parents=True, exist_ok=True)
        return d
    d = wr.subject_dir(batch_dir_name(batch_no), subject)
    d.mkdir(parents=True, exist_ok=True)
    return d


def copy_root_for(args) -> Path | None:
    r"""源文件副本存哪：显式 `--copy-root` 优先；否则 `<工作区>\in`。"""
    explicit = getattr(args, "copy_root", None)
    if explicit:
        return Path(explicit).expanduser()
    wr = getattr(args, "_workroot", None)
    if wr is None or not is_workroot(wr.path):
        try:
            wr = open_for_command(args)
        except WorkrootError:
            return None
    if not is_workroot(wr.path):
        return None
    wr.ensure()
    return wr.inbox


def cmd_work(args):  # pragma: no cover - exercised by selftest
    from .common import Result, out_dir, unique_path, write_text

    res = Result("work")

    # ---- 分诊（临时编程：共性问题 / 特殊问题）----------------------------
    if getattr(args, "triage", False):
        info = triage(args)
        text = render_triage(info)
        res.data.update({"action": "triage", **info})
        if info.get("留痕失败"):
            res.warn("留痕没写进事件链：%s" % info["留痕失败"])
        if info.get("kind") == "special":
            res.warn("已复制到 %s 并写了 FORK.md —— **D:\\DSH\\tool 一个字都没改**；"
                     "以后这份活儿要跑副本里的 office.py" % info.get("fork_path"))
        else:
            res.warn("判为共性问题 = **工具升级**；请到能改代码的会话里改 %s，"
                     "改完跑六套自测" % ROOT_OF_TOOLKIT)
        p = write_text(unique_path(out_dir(getattr(args, "out", None), "work")
                                   / "分诊结论.md"), text)
        res.add_artifact(p, "临时编程分诊结论（共性问题 / 特殊问题）")
        res.data["report"] = text
        return res

    # ---- 建议一个工作区（给使用者审核用）--------------------------------
    if getattr(args, "suggest", False):
        info = suggest(for_what=getattr(args, "for_what", None))
        text = render_suggest(info)
        res.data.update({"action": "suggest", **info})
        res.warn("工作区**还没定**。请把下面这个建议给使用者确认（他回一句「可以」就开工）：\n"
                 "  %s" % info["建议完整路径"])
        res.data["report"] = text
        return res

    # ---- 动手前那句话（原样念给使用者）----------------------------------
    if getattr(args, "brief", False):
        text = brief()
        res.data.update({"action": "brief", "report": text})
        return res

    if getattr(args, "init", False):
        at = getattr(args, "at", None) or Path.cwd()
        info = init_workroot(at, note=getattr(args, "note", None))
        res.data.update({"action": "init", **info})
        text = render_work("init", info)
        res.warn("工作区已建在 %s —— 以后在这个文件夹里跑 db-* 命令，"
                 "库、副本、产物、日志都会落在它下面" % info["workroot"])
    else:
        wr, db = resolve(work=getattr(args, "work", None))
        wr.ensure()
        info = wr.describe()
        info["库大小"] = db.stat().st_size if db.exists() else 0
        info["子目录"] = {s: str(wr.path / s) for s in SUBDIRS}
        info.update(session_report(wr))
        res.data.update({"action": "status", **info})
        text = render_work("status", info)
        if not db.exists():
            res.warn("工作区在 %s，但库里还没有东西（还没跑过 db-ingest）" % wr.path)
        if info.get("当前目录在工作区内") is False:
            res.warn("⚠️ 当前目录 %s **不在**工作区 %s 里面——"
                     "确认一下你是不是想用另一个工作区？" % (info["当前目录"], wr.path))
    if getattr(args, "out", None):
        p = write_text(unique_path(out_dir(args.out, "work") / "workroot_report.md"), text)
        res.add_artifact(p, "工作区报告")
    res.data["report"] = text
    return res


def render_work(kind: str, info: dict) -> str:
    L: list[str] = []
    A = L.append
    if kind == "init":
        A("# 工作区已就绪")
        A("")
        A("**工作区 = `%s`**" % info["workroot"])
        A("")
        A("| 位置 | 用途 |")
        A("| --- | --- |")
        A("| `in\\<批次号>\\` | 源文件的整份副本（只读） |")
        A("| `out\\<批次号>\\<主体名>\\` | 填好的目标文件 |")
        A("| `out\\<批次号>\\_报告\\<命令>\\` | 人看的报告 |")
        A("| `db\\workflow.db` | 数据库（SCHEMA v2） |")
        A("| `logs\\` | 运行日志（按天滚动） |")
        A("| `work\\<批次号>\\` | 临时区（可清理） |")
        A("")
        A("标记文件 `%s` 已写好——**在这个文件夹（或它的子目录）里跑命令，"
          "命令会自动认领这个工作区**，不用每次给路径。" % MARKER_NAME)
        if info.get("created_dirs"):
            A("")
            A("本次新建的子目录：%s" % "、".join(info["created_dirs"]))
    else:
        A("# 当前工作区")
        A("")
        A("| 项 | 值 |")
        A("| --- | --- |")
        for k, v in info.items():
            if k in ("子目录", "能读到的 DSH 环境变量"):
                continue
            A("| %s | %s |" % (k, v))
        A("")
        A("## 目录")
        A("")
        A("| 子目录 | 位置 |")
        A("| --- | --- |")
        for k, v in (info.get("子目录") or {}).items():
            A("| `%s\\` | %s |" % (k, v))
        A("")
        A("## 🔴 动手前先对齐三件事（**工具 / 工作区 / 权限**）")
        A("")
        A("| # | 要对齐什么 | 现在的情况 |")
        A("| --- | --- | --- |")
        A("| 1 | **工具**在哪 | `%s`（只读即可运行；**要改它**才需要写权限）"
          % Path(__file__).resolve().parent.parent)
        A("| 2 | **工作区**在哪 | office 工作区 = `%s`；当前目录 = `%s`；"
          "**当前目录在工作区内 = %s**"
          % (info.get("工作区"), info.get("当前目录"), info.get("当前目录在工作区内")))
        A("| 3 | **权限**够不够 | **本工具读不到权限档位**（DSH 不传）。"
          "**必须由代理在对话里说明**：读工具包 + 写工作区 → 至少「工作区内修改」；"
          "**要改工具包本身或写到工作区之外 → 需要「完全权限」，且应当单独征求同意** |")
        A("")
        A("> 📌 **每个工作区自己带一个标记文件**；命令从当前目录往上找到它就认。")
        A("> 找不到就**报错**，**不会**把库悄悄建在当前目录（免得文件不知道跑哪去了）。")
        A("")
        A("⚠️ **诚实边界**：DSH 目前只往子进程传 **4 个环境变量**（%s），"
          "**不传会话工作区、也不传权限档位**。所以\"DSH 选的那个工作区\"与"
          "\"office 的工作区\"是不是同一个，**本工具判不了**——"
          "只能报出当前目录与工作区的关系，剩下那句必须由代理说。"
          % "、".join(info.get("能读到的 DSH 环境变量") or []))
    return "\n".join(L) + "\n"
