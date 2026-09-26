"""`db-absorb` —— 从材料里读出候选键值对，**摊成一张人能审的表**（T-27）。

它解决的是哪一段
----------------
现在这段是**断的**：材料 →（谁来读？）→ `profile.json` → 入库。
写 JSON 的一直是代理，**而使用者看不懂 JSON —— 于是"用户审核"在第一步就失效了**：
你看不懂的东西，没法审。

`db-absorb` 把这一步变成**一张中文表，直接打在对话里**：

```
  #  材料里的原话（第几行）          我读出来的        这条是谁的        收不收
  1  借款人：广东众森实业发展有限公司   借款人名称       广东众森实业…     ✅ 建议收
  4  联系电话：13712381510          联系电话         ⚠️ 认不出是谁的   🔴 要你指明
  —  （材料里没写）                  用信用途         —                ⬜ 留空待你填
```

**三条设计规矩**（主人第 24 轮已定）：

1. **每一行都带"材料里的原话"和行号**——你审的是"我有没有看错"，就得把原文摆在你眼前；
2. **永远不给使用者看 JSON**——JSON 只是给工具读回去的中间物；
3. **逐条四选**：收下 / 我改一下 / 这条不要 / 先空着；**外加"我补一条"**。

⚠️ **诚实边界**：这里的抽取是**机械的**——扫 `标签：值`，不做理解。
所以它**只对"摘要式"材料管用**（一行一个标签），合同正文那种它读不出来。
**真正难的材料仍由大模型读懂后填进同一张表**——表是同一个，审的通道是同一条。
字段键名也是**从材料上抄的**，还没和 [18 字段册] 对齐（那是待办 T-4）。
"""
from __future__ import annotations

import json
import re
from pathlib import Path
from typing import Any

from . import doc_read
from .common import OfficeKitError, Result, unique_path, write_text

#: 一行里"标签：值"的样子。标签短、值跟在后面。
LINE_RE = re.compile(r"^\s*([\u4e00-\u9fffA-Za-z0-9（）()、·*]{1,24})\s*[:：]\s*(.+?)\s*$")
#: 这一行在说"谁是借款人/保证人/法定代表人"
SUBJECT_WORDS = {
    "借款人名称": "借款人", "保证人名称": "保证人",
    "借款人": "借款人", "客户": "借款人", "申请人": "借款人", "授信申请人": "借款人",
    "用信人": "借款人",
    "保证人": "保证人", "担保人": "保证人", "保证方": "保证人",
    "法定代表人": "法定代表人", "法人代表": "法定代表人", "法人": "法定代表人",
}
#: 标签 → 标准键名（**只放我们有把握的**；其余照抄标签并标 ⚠️）
LABEL_ALIASES = {
    "借款人": "借款人名称", "客户": "借款人名称", "申请人": "借款人名称",
    "授信申请人": "借款人名称", "用信人": "借款人名称",
    "保证人": "保证人名称", "担保人": "保证人名称", "保证方": "保证人名称",
    "统一社会信用代码": "统一社会信用代码", "行内编号": "行内编号",
    "联系电话": "联系电话", "电话": "联系电话", "联系方式": "联系电话",
    "收款账号": "收款账号", "账号": "收款账号", "账户": "收款账号",
    "开户行": "开户行", "开户银行": "开户行",
    "用信金额": "用信金额", "用信金额（元）": "用信金额",
    "用信用途": "用信用途", "用途": "用信用途", "资金用途": "用信用途",
    "证件类型": "证件类型", "证件种类": "证件类型",
    "证件号码": "证件号码", "证件号": "证件号码",
    "开户行及账号": "开户行及账号", "开户银行及账号": "开户行及账号",
    "保证金额": "保证金额", "保证期限": "保证期限", "期限": "期限",
}
#: 这些标签的**值本身是一个主体名**（不是普通字段）
SUBJECT_LABELS = {"借款人名称", "保证人名称", "借款人法定代表人", "保证人法定代表人",
                  "个人保证人姓名"}

#: 只接受有限、可解释的编号主体标签；避免把任意以“保证人”开头的字段当作主体。
NUMBERED_GUARANTOR_RE = re.compile(
    r"^(保证人|担保人|保证方)(?:[（(]?([0-9一二三四五六七八九十]+)[）)]?)?(?:名称)?$")
LEGAL_REP_LABELS = {"法定代表人", "法人代表", "法人", "法定代表人姓名", "法人姓名"}
#: 值为证件号码时，只有标签本身明确类型才补出证件类型；不把“证件号码”猜成身份证。
CERTIFICATE_NUMBER_LABELS = {
    "证件号码": None, "证件号": None,
    "身份证号": "居民身份证", "身份证号码": "居民身份证",
    "居民身份证号": "居民身份证", "居民身份证号码": "居民身份证",
    "护照号": "护照", "护照号码": "护照",
    "港澳居民来往内地通行证号": "港澳居民来往内地通行证",
}
for _certificate_name in ("居民身份证", "护照", "台湾居民来往大陆通行证",
                          "港澳居民来往内地通行证", "台湾居民居住证", "港澳居民居住证",
                          "外国人永久居留身份证", "军官证", "士兵证", "警官证"):
    for _suffix in ('', '号', '号码', '编号'):
        CERTIFICATE_NUMBER_LABELS[_certificate_name + _suffix] = _certificate_name
CERTIFICATE_NUMBER_LABELS['身份证'] = '居民身份证'
CERTIFICATE_TYPE_LABELS = {"证件类型", "证件种类"}


def _known_inline_labels() -> tuple[str, ...]:
    """返回可在同一句中再次出现的明确字段标签。

    这里故意只允许已知标签。比如“开户地址：中国：合成支行”中的普通冒号，
    不能被猜成另一个字段的边界。
    """
    labels = set(LABEL_ALIASES) | set(CERTIFICATE_NUMBER_LABELS) | set(CERTIFICATE_TYPE_LABELS)
    return tuple(sorted(labels, key=len, reverse=True))


INLINE_LABEL_RE = re.compile(
    r"\s+(%s)\s*[:：]" % "|".join(re.escape(label) for label in _known_inline_labels()))

# 仅把可能影响填写的叙述留痕。合同标题、普通说明和没有业务数值的段落不在此列。
NARRATIVE_FACT_RE = re.compile(
    r"(?:金额|借款|贷款|用信|授信|人民币).{0,24}?\d[\d,]*(?:\.\d+)?\s*(?:元|万元|人民币)"
    r"|(?:账号|账户|卡号).{0,12}?\d{6,}"
    r"|(?:身份证|护照|通行证|证件).{0,16}?[A-Za-z0-9-]{6,}"
)
# 只接受这个完整、无歧义的句式；不换算单位，也不从含多个金额的合同句子猜字段。
STRICT_LOAN_AMOUNT_RE = re.compile(r"^本次借款金额为人民币(\d+(?:\.\d+)?)元[。.]?$")


def _split_explicit_labels(line: int, text: str) -> list[tuple[int, str, str]]:
    """把一行中的多个已知“标签：值”拆开，所有候选仍指向完整原句。"""
    first = LINE_RE.match(text)
    if not first:
        return [(line, text, text)]
    entries: list[tuple[str, str]] = []
    label, remainder = first.group(1).strip(), first.group(2).strip()
    while True:
        later = INLINE_LABEL_RE.search(remainder)
        if not later:
            entries.append((label, remainder))
            break
        value = remainder[:later.start()].strip()
        if value:
            entries.append((label, value))
        label = later.group(1)
        remainder = remainder[later.end():].strip()
    return [(line, "%s：%s" % (entry_label, value), text)
            for entry_label, value in entries if value]


def _narrative_issue(line: int, text: str) -> dict[str, Any] | None:
    if not NARRATIVE_FACT_RE.search(text):
        return None
    return {
        "line": line,
        "quote": text,
        "reason": "叙述中含金额、账号或证件相关信息，但未按“标签：值”解析；请确认是否需要补充字段。",
        "kind": "unparsed_business_fact",
    }


def read_lines(path: Path) -> list[tuple[int, str]]:
    """把材料读成 `(行号, 文字)`。**行号是唯一的"出处"**，绝不能丢。

    表格里的"左边一格标签、右边一格值"会被合成为 `标签：值` 这样一行
    ——因为对人来说，那就是一行。
    """
    suffix = path.suffix.lower()
    lines: list[tuple[int, str]] = []
    if suffix == ".docx":
        data = doc_read.read_docx(path)
        for block in data["blocks"]:
            if block["type"] in ("paragraph", "list_item"):
                text = str(block.get("text") or "").strip()
                if text:
                    lines.append((len(lines), text))
            elif block["type"] == "table":
                for row in ([block.get("header")] if block.get("header") else []) + \
                        list(block.get("data") or []):
                    cells = [str(c or "").strip() for c in row]
                    i = 0
                    while i < len(cells):
                        label = cells[i].rstrip("：: ")
                        if label and i + 1 < len(cells) and cells[i + 1]:
                            lines.append((len(lines), "%s：%s" % (label, cells[i + 1])))
                            i += 2
                        else:
                            if cells[i]:
                                lines.append((len(lines), cells[i]))
                            i += 1
    elif suffix in (".xlsx", ".xlsm", ".xls"):
        for block in doc_read.read_workbook(path)["sheets"]:
            header = block.get("header") or []
            for row in block.get("preview") or []:
                cells = [str(c or "").strip() for c in row]
                for i, cell in enumerate(cells):
                    if not cell:
                        continue
                    label = cell.rstrip("：: ")
                    nxt = cells[i + 1] if i + 1 < len(cells) else ""
                    if nxt and len(label) <= 24:
                        lines.append((len(lines), "%s：%s" % (label, nxt)))
                    else:
                        lines.append((len(lines), cell))
            for i, cell in enumerate(header):
                if cell:
                    lines.append((len(lines), str(cell)))
    elif suffix in (".md", ".txt", ".csv", ".tsv"):
        text = path.read_text(encoding="utf-8", errors="replace")
        for raw in text.splitlines():
            if raw.strip():
                lines.append((len(lines), raw.strip()))
    else:
        raise OfficeKitError(
            "这个格式我读不了：%s。先用 `office.py extract <文件>` 转成 markdown/txt 再试。"
            % suffix)
    return lines


def absorb_with_diagnostics(path: Path) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    """读取候选行及未按标签解析的有限业务叙述。

    ``issues`` 的每项都有 ``line``、``quote``、``reason``，供工作流在来源
    快照中保存并在需要时交给使用者确认。它不把普通说明文字升级为问题。
    """
    rows: list[dict[str, Any]] = []
    issues: list[dict[str, Any]] = []
    role_owner: dict[str, str] = {}       # 角色 → 最后一个可回退主体（仅无编号角色）
    numbered_guarantors: dict[str, str] = {}
    last_company_role: str | None = None  # 最近一次出现的主体角色（借款人或保证人）
    last_company_owner: str | None = None # 最近明确主体；不能被“保证人”角色桶覆盖
    last_company_label: str | None = None # 原文角色标签，供提问卡清楚说明编号保证人
    last_natural_person: dict[str, str] | None = None
    section_kind: str | None = None
    section_context: str | None = None
    pending_guarantee_titles: list[int] = []

    def add_row(**row: Any) -> None:
        row["n"] = len(rows) + 1
        rows.append(row)

    def company_role(raw_label: str) -> str | None:
        direct = SUBJECT_WORDS.get(raw_label)
        if direct in ("借款人", "保证人"):
            return direct
        return "保证人" if NUMBERED_GUARANTOR_RE.fullmatch(raw_label) else None

    def guarantor_number(raw_label: str) -> str | None:
        match = NUMBERED_GUARANTOR_RE.fullmatch(raw_label)
        return match.group(2) if match and match.group(2) else None

    def owner_for_role(role: str | None) -> str | None:
        """Return the current explicit owner before a generic role fallback.

        Several guarantors share the same role.  ``role_owner['保证人']`` is
        therefore never enough to decide who owns the fields following a
        numbered guarantor declaration.
        """
        if role and role == last_company_role and last_company_owner:
            return last_company_owner
        return role_owner.get(role) if role else None

    for n, text, quote in (part for original_line, original_text in read_lines(path)
                           for part in _split_explicit_labels(original_line, original_text)):
        m = LINE_RE.match(text)
        section_label = m.group(1) if m else text
        # 合同标题是有限的段落边界信号。保证合同不能沿用上一个主体；贷款/授信
        # 合同可回到材料已明确的借款人。标题本身不当作一个待入库字段。
        if "保证合同" in section_label:
            # A contract number often appears in the heading immediately
            # before the form declares its guarantor.  Keep only that title
            # pending for the next explicit guarantor; do not let ordinary
            # addresses/phones cross this contract boundary.
            pending_guarantee_titles.clear()
            last_company_role = None
            last_company_owner = None
            last_company_label = None
            last_natural_person = None
            section_kind = "保证合同"
            section_context = "当前材料段为保证合同，尚未识别保证主体"
        elif "流动资金贷款合同" in section_label or "借款合同" in section_label or "贷款合同" in section_label:
            last_company_role = "借款人" if role_owner.get("借款人") else None
            last_company_owner = role_owner.get("借款人")
            last_company_label = "借款人" if last_company_owner else None
            last_natural_person = None
            section_kind = "贷款合同"
            section_context = "当前材料段为贷款合同"
        elif "授信合同" in section_label or "授信协议" in section_label:
            last_company_role = "借款人" if role_owner.get("借款人") else None
            last_company_owner = role_owner.get("借款人")
            last_company_label = "借款人" if last_company_owner else None
            last_natural_person = None
            section_kind = "授信合同"
            section_context = "当前材料段为授信合同"
        if not m:
            amount = STRICT_LOAN_AMOUNT_RE.fullmatch(text)
            if amount:
                owner = owner_for_role(last_company_role)
                context = ("位于最近明确的%s“%s”段内；叙述金额候选需确认" %
                           (last_company_role, owner) if owner else
                           "叙述金额候选未能确认归属主体，需确认")
                if section_context:
                    context = "%s；%s" % (context, section_context)
                add_row(line=n, label="本次借款金额", key="借款金额",
                        value="%s元" % amount.group(1), quote=quote,
                        entity_name=owner, kind="字段", role=None,
                        owner_kind="subject" if owner else "unknown", context=context,
                        confidence=0.7 if owner else 0.5,
                        needs="默认" if owner else "指明主体", assumed=bool(owner),
                        known_key=True, ownership_group=None, certificate_type=None)
                continue
            issue = _narrative_issue(n, quote)
            if issue and issue not in issues:
                issues.append(issue)
            continue
        raw_label, value = m.group(1).strip(), m.group(2).strip()
        if not value:
            continue
        if len(value) > 80:
            if (raw_label in LABEL_ALIASES or raw_label in CERTIFICATE_NUMBER_LABELS or
                    raw_label in CERTIFICATE_TYPE_LABELS or NARRATIVE_FACT_RE.search(text)):
                issues.append({
                    "line": n, "quote": quote,
                    "reason": "明确业务字段的值超过 80 个字符，未自动截断或解析；请确认需要保留的字段值。",
                    "kind": "unparsed_business_fact",
                })
            continue
        role = company_role(raw_label)
        if role:
            # 编号保证人仍是“保证人”角色；原标签留在 label/quote 中用于回溯。
            role_owner[role] = value
            number = guarantor_number(raw_label) if role == "保证人" else None
            if number:
                numbered_guarantors[number] = value
            last_company_role = role
            last_company_owner = value
            last_company_label = raw_label
            last_natural_person = None
            if section_kind == "保证合同" and role == "保证人":
                section_context = "当前保证合同段已明确保证主体"
                for row_index in pending_guarantee_titles:
                    pending = rows[row_index]
                    pending.update(entity_name=value, owner_kind="subject",
                                   context=("保证合同标题在第%d行，后续材料明确声明保证人“%s”"
                                            % (pending["line"], value)),
                                   confidence=1.0, needs="收", assumed=False)
                pending_guarantee_titles.clear()
            add_row(line=n, label=raw_label, key=role + "名称", value=value, quote=quote,
                    entity_name=value, kind="主体", role=role, owner_kind="subject",
                    context="材料明确声明该主体为%s" % role, confidence=1.0, needs="收",
                    assumed=False, known_key=True)
            continue

        explicit_legal_role = ("借款人" if raw_label.startswith("借款人") and
                               raw_label.endswith(("法定代表人", "法人代表", "法人")) else
                               "保证人" if raw_label.startswith("保证人") and
                               raw_label.endswith(("法定代表人", "法人代表", "法人")) else None)
        if raw_label in LEGAL_REP_LABELS or explicit_legal_role:
            legal_role = explicit_legal_role or last_company_role
            company = owner_for_role(legal_role)
            key = (legal_role + "法定代表人") if legal_role else "法定代表人"
            context = ("材料明确写明其为%s“%s”的法定代表人" % (legal_role, company)
                       if company else "材料写明法定代表人姓名，但未说明关联主体")
            last_natural_person = {"name": value, "company": company or "",
                                   "company_role": legal_role or ""}
            add_row(line=n, label=raw_label, key=key, value=value, quote=quote,
                    entity_name=company, kind="主体关系", role=None,
                    owner_kind="subject" if company else "unknown", person_name=value,
                    context=context, confidence=1.0 if company else 0.6,
                    needs="收" if company else "指明主体", assumed=False,
                    known_key=bool(company))
            continue

        certificate_type = CERTIFICATE_NUMBER_LABELS.get(raw_label)
        is_certificate_number = raw_label in CERTIFICATE_NUMBER_LABELS
        is_certificate_type = raw_label in CERTIFICATE_TYPE_LABELS
        explicit_role = next((candidate for candidate in ("借款人", "保证人")
                              if raw_label.startswith(candidate)), None)
        key = ("证件号码" if is_certificate_number else
               "证件类型" if is_certificate_type else LABEL_ALIASES.get(raw_label, raw_label))
        if raw_label == "期限" and section_kind in ("贷款合同", "授信合同", "保证合同"):
            key = ("贷款期限" if section_kind == "贷款合同" else
                   "授信期限" if section_kind == "授信合同" else "保证期限")
        owner = None
        owner_kind = "unknown"
        context = "材料没有足以确认归属的主体上下文"
        assumed = False
        if is_certificate_number or is_certificate_type:
            if last_natural_person:
                owner = last_natural_person["name"]
                owner_kind = "person"
                relation = last_natural_person["company"] or "未标明主体"
                context = "紧随法定代表人“%s”（关联主体：%s）的证件信息" % (owner, relation)
                # The immediately preceding legal-representative declaration
                # explicitly links this person to a named company.  Treat the
                # following certificate as source evidence, not an inferred
                # ownership question.  A bare legal-representative name still
                # has no company relation and remains for confirmation.
                assumed = not bool(last_natural_person["company"])
            elif last_company_role and owner_for_role(last_company_role):
                owner = owner_for_role(last_company_role)
                owner_kind = "subject"
                context = "紧随最近明确的%s“%s”段，未见新的法定代表人" % (last_company_role, owner)
                # A named borrower/guarantor (including a natural-person
                # guarantor) is direct source ownership for the immediately
                # following certificate.  Contract headings reset this state,
                # so it cannot leak into the next agreement section.
                assumed = not bool(last_company_owner)
        else:
            target_role = explicit_role or last_company_role
            if target_role and owner_for_role(target_role):
                owner = owner_for_role(target_role)
                owner_kind = "subject"
                context = ("标签明确指向%s" % target_role if explicit_role else
                           "位于最近明确的%s“%s”段内" % (target_role, owner))
                # A field following an explicit borrower/guarantor declaration
                # is direct source evidence, not a repeated ownership question.
                # Contract headings clear last_company_owner, so cross-section
                # and no-subject material still remain for confirmation.
                assumed = not bool(explicit_role)
                if last_company_role in ("借款人", "保证人") and last_company_owner:
                    assumed = False
        if section_context:
            context = "%s；%s" % (context, section_context)
        ownership_group = ("certificate:%s\x1f%s" % (quote, owner or "")
                           if is_certificate_number and certificate_type else None)

        add_row(line=n, label=raw_label, key=key, value=value, quote=quote,
                entity_name=owner, kind="字段", role=None,
                owner_kind=owner_kind, context=context, confidence=0.95 if owner else 0.6,
                needs="收" if owner and not assumed else "默认" if owner else "指明主体",
                assumed=assumed, known_key=(raw_label in LABEL_ALIASES or
                                             is_certificate_number or is_certificate_type),
                ownership_group=ownership_group, certificate_type=certificate_type,
                owner_heading=last_company_label if owner else None)
        if (section_kind == "保证合同" and "保证合同" in raw_label and
                owner is None and not is_certificate_number and not is_certificate_type):
            pending_guarantee_titles.append(len(rows) - 1)
        if is_certificate_number and certificate_type:
            # 类型来自“身份证号”等原标签本身，属于材料明示，不是默认身份证。
            add_row(line=n, label=raw_label, key="证件类型", value=certificate_type, quote=quote,
                    entity_name=owner, kind="字段", role=None,
                    owner_kind=owner_kind, context=context + "；证件类型由原标签明确",
                    confidence=0.95 if owner else 0.6,
                    needs="收" if owner and not assumed else "默认" if owner else "指明主体",
                    assumed=assumed, known_key=True, ownership_group=ownership_group,
                    certificate_type=certificate_type,
                    owner_heading=last_company_label if owner else None)
    return rows, issues


def absorb(path: Path) -> list[dict[str, Any]]:
    """兼容旧调用：只返回候选行。需要来源诊断时用 ``absorb_with_diagnostics``。"""
    rows, _issues = absorb_with_diagnostics(path)
    return rows


def render(rows: list[dict], source: str, counts: dict) -> str:
    L: list[str] = []
    A = L.append
    A("## 源材料解析记录（待确认项由 DSH 弹窗提问）")
    A("")
    A("来源：`%s`" % source)
    A("")
    if not rows:
        A("（没读到任何 `标签：值` 形式的行。这份材料可能不是摘要式的，"
          "需要我读懂之后再填进同一张表。）")
        return "\n".join(L) + "\n"
    A("| # | 材料里的原话（第几行） | 我读出来的 | 这条是谁的 | 收不收 |")
    A("| --- | --- | --- | --- | --- |")
    for r in rows:
        if r.get("assumed"):
            who = "🔸 默认**%s**，不对就说" % r["entity_name"]
        elif r["entity_name"]:
            who = r["entity_name"]
        else:
            who = "⚠️ **认不出是谁的**"
        mark = {"收": "✅ 建议收", "默认": "🔸 默认（请扫一眼）"}.get(
            r["needs"], "🔴 **要你指明**")
        key = r["key"] if r["known_key"] else "%s ⚠️" % r["key"]
        where = "（第%d行）" % r["line"] if isinstance(r.get("line"), int) else "（你补的）"
        A("| %d | %s%s | `%s`<br>= %s | %s | %s |"
          % (r["n"], _q(r["quote"]), where, key, _v(r["value"]), who, mark))
    A("")
    A("**这份文件仅供事后查阅。**归属不明的项目在 DSH 提问卡片中选择建议、暂不收录或自定义主体；不在这份表里回答。")
    A("")
    A("> ⚠️ **标 ⚠️ 的键名是我从材料上抄的**，还没和字段册对齐（待办 T-4）。")
    A("> 抄错了我认，但名字**照原样**给你看，不替你改。")
    if counts.get("assumed"):
        A("")
        A("🔸 **有 %d 条我按材料上下文默认归属主体**（材料里没写主语），已列入弹窗待确认项。" % counts["assumed"])
    if counts.get("needs"):
        A("")
        A("🔴 **有 %d 条连主体都还没有**（材料里没出现任何主体名，或根本没读到），已列入弹窗待确认项。" % counts["needs"])
    return "\n".join(L) + "\n"


def source_questions(rows: list[dict]) -> list[dict]:
    """把归属不明确的源字段交给 DSH 原生提问卡片，不让用户审长表。"""
    candidates: list[dict[str, str]] = []

    def add_candidate(name: str | None, kind: str, role: str | None,
                      context: str | None) -> None:
        if not name or any(c["name"] == name and c["kind"] == kind for c in candidates):
            return
        candidates.append({"name": name, "kind": kind, "role": role or "未说明角色",
                           "context": context or "材料中出现该主体"})

    for row in rows:
        if row.get("kind") == "主体" and row.get("owner_kind") == "subject":
            add_candidate(row.get("entity_name"), "主体", row.get("role"), row.get("context"))
        if row.get("person_name"):
            add_candidate(row["person_name"], "个人", "法定代表人", row.get("context"))
        if row.get("owner_kind") == "person":
            add_candidate(row.get("entity_name"), "个人", "法定代表人", row.get("context"))
    certificate_groups: dict[tuple[str | None, str], list[dict[str, Any]]] = {}
    for row in rows:
        if row.get("ownership_group"):
            certificate_groups.setdefault((row.get('_source'), row["ownership_group"]), []).append(row)
    asked_groups: set[tuple[str | None, str]] = set()
    questions = []
    for r in rows:
        if r.get("dropped") or not (r.get("assumed") or r.get("needs") == "指明主体"):
            continue
        ownership_group = r.get("ownership_group")
        group_key = (r.get('_source'), ownership_group)
        if ownership_group and group_key in asked_groups:
            continue
        if ownership_group:
            asked_groups.add(group_key)
        suggested = str(r.get("entity_name") or "")
        options = []
        if suggested and r.get("context"):
            options.append({"label": "归属：%s（推荐）" % suggested,
                            "description": "%s；尚未由你确认。" % r["context"]})
        else:
            options.append({"label": "暂不收录", "description": "没有可靠归属依据时先留空，也可从下面选择核实后的主体。"})
        for candidate in candidates:
            if candidate["name"] == suggested:
                continue
            options.append({"label": "归属：%s" % candidate["name"],
                            "description": "%s，角色：%s。%s" %
                                           (candidate["kind"], candidate["role"],
                                            candidate["context"])})
        if not any(option['label'] == '暂不收录' for option in options):
            options.append({"label": "暂不收录", "description": "这一条不入库，后续需要时再处理。"})
        source_name = r.get("_source") or r.get("source") or "当前材料"
        grouped_rows = certificate_groups.get(group_key, [r]) if ownership_group else [r]
        number = next((row["value"] for row in grouped_rows if row["key"] == "证件号码"), None)
        certificate_type = next((row["value"] for row in grouped_rows if row["key"] == "证件类型"),
                                r.get("certificate_type"))
        if number:
            detail = "来源：%s。证件号码：%s。证件类型：%s。原文：%s。" % (
                source_name, number, certificate_type or "材料未写明", r["quote"])
        else:
            detail = "来源：%s。字段：%s。原文：%s。" % (source_name, r["key"], r["quote"])
        if suggested and r.get("context"):
            detail += "建议归属：%s。依据：%s。请确认或修改。" % (suggested, r["context"])
        else:
            detail += "材料未给出足以确认归属的上下文；请指定公司或个人，或暂不收录。"
        heading = r.get("owner_heading")
        owner_title = ("%s %s · " % (heading, suggested)
                       if heading and suggested else "")
        questions.append({
            "id": "source-%s" % r["n"], "header": "%s材料归属 · %s" % (owner_title, r["key"]),
            "question": detail,
            "options": options,
        })
    return questions


def _q(s: str) -> str:
    t = str(s).replace("|", "\\|")
    return ("`%s`" % t) if len(t) <= 34 else ("`%s…`" % t[:33])


def _v(s: str) -> str:
    t = str(s).replace("|", "\\|")
    return t if len(t) <= 24 else t[:23] + "…"


def apply_answers(rows: list[dict], *, assign: list[str] | None = None,
                  setv: list[str] | None = None, add: list[str] | None = None,
                  drop: list[str] | None = None) -> tuple[list[dict], list[str]]:
    """把使用者的回答落到草稿上。**对不上的编号直接报错，不忽略。**"""
    by_n = {r["n"]: r for r in rows}
    problems: list[str] = []
    out = [dict(r) for r in rows]

    def _n(token: str) -> int | None:
        try:
            return int(str(token).split("=")[0].strip())
        except ValueError:
            return None

    for token in assign or []:
        if "=" not in str(token):
            problems.append("--assign 要写成 编号=主体名，例如 --assign 4=广东众森实业发展有限公司")
            continue
        num, who = str(token).split("=", 1)
        n = _n(num)
        if n not in by_n:
            problems.append("没有第 %s 条" % num)
            continue
        for r in out:
            if r["n"] == n:
                r["entity_name"] = who.strip()
                r["needs"] = "收"
                r["assigned_by_user"] = True
    for token in setv or []:
        if "=" not in str(token):
            problems.append("--set 要写成 编号=值，例如 --set 2=顾红军")
            continue
        num, val = str(token).split("=", 1)
        n = _n(num)
        if n not in by_n:
            problems.append("没有第 %s 条" % num)
            continue
        for r in out:
            if r["n"] == n:
                r["value"] = val.strip()
                r["user_edited"] = True
    for token in drop or []:
        n = _n(token)
        if n not in by_n:
            problems.append("没有第 %s 条" % token)
            continue
        for r in out:
            if r["n"] == n:
                r["dropped"] = True
    for token in add or []:
        if "=" not in str(token):
            problems.append("--add 要写成 键=值，例如 --add 用信用途=采购原材料")
            continue
        key, val = str(token).split("=", 1)
        out.append({"n": len(out) + 1, "line": None, "label": key.strip(),
                    "key": key.strip(), "value": val.strip(), "quote": "（你补的，材料里没有）",
                    "entity_name": None, "kind": "补充", "confidence": 1.0,
                    "needs": "收", "known_key": True, "from_user": True})
    return out, problems


def to_profile(rows: list[dict], *, default_entity: str | None = None) -> dict[str, Any]:
    """把草稿变成 `db-ingest` 吃得下的字段字典。

    * 丢弃的行**不进字典**；
    * 使用者补的行标 `source_kind: user`（**不无中生有**，红线二）；
    * 出处一律是 `"<材料名> 第N行"` 或 `"你补充的"`。
    """
    prof: dict[str, Any] = {}
    for r in rows:
        if r.get("dropped"):
            continue
        entry: dict[str, Any] = {"value": r.get("value"),
                                 "source": r.get("quote") if r.get("line") is None
                                 else "%s" % r["quote"]}
        if r.get("from_user"):
            entry["source"] = "你补充的"
            entry["source_kind"] = "user"
            entry["note"] = "使用者补的（材料里没有）"
        if r.get("entity_name"):
            entry["entity_name"] = r["entity_name"]
        if r.get("needs") == "指明主体" and not r.get("entity_name"):
            entry["note"] = "⚠️ 认不出是谁的（未指明主体）"
        prof[r["key"]] = entry
    if default_entity:
        prof["_默认主体"] = default_entity
    return prof


def cmd_db_absorb(args) -> Result:
    from . import workroot as WR

    res = Result("db-absorb")
    try:
        wr = WR.open_for_command(args)
    except WR.WorkrootError as exc:
        raise OfficeKitError(str(exc)) from exc

    files = [Path(p) for p in (args.input if isinstance(args.input, list) else [args.input])]
    src = files[0]
    if not src.exists():
        store_close = getattr(wr, "log", None)
        raise OfficeKitError("材料不存在：%s" % src)

    rows = absorb(src)
    rows, problems = apply_answers(
        rows, assign=getattr(args, "assign", None), setv=getattr(args, "set", None),
        add=getattr(args, "add", None), drop=getattr(args, "drop", None))

    counts = {"all": len(rows),
              "assumed": sum(1 for r in rows if r.get("assumed") and not r.get("dropped")),
              "needs": sum(1 for r in rows
                           if r["needs"] == "指明主体" and not r["entity_name"]),
              "dropped": sum(1 for r in rows if r.get("dropped")),
              "added": sum(1 for r in rows if r.get("from_user"))}
    text = render(rows, src.name, counts)

    out = WR.report_dir_for(args, "db-absorb", getattr(args, "batch", None))
    p_md = write_text(unique_path(out / "profile_draft.md"), text)
    prof = to_profile(rows, default_entity=getattr(args, "default_entity", None))
    p_json = write_text(unique_path(out / "profile_draft.json"),
                        json.dumps(prof, ensure_ascii=False, indent=2))
    res.add_artifact(p_md, "源材料解析记录（留存备查，不用它索取确认）")
    res.add_artifact(p_json, "同一份草稿（入库时用 --profile 传回去）")

    res.data.update({"source": str(src), "rows": rows, "counts": counts,
                     "questions": source_questions(rows),
                     "profile_file": str(p_json), "problems": problems, "report": text})
    if problems:
        res.warn("有 %d 处没听懂（%s）" % (len(problems), "；".join(problems[:2])))
    if counts["assumed"]:
        res.warn("🔸 有 %d 条我**按材料上下文默认归属主体**（材料里没写主语）——"
                 "请用 data.questions 调 DSH 的 ask_user_question 弹窗确认，不用 Markdown 表格确认"
                 % counts["assumed"])
    if counts["needs"]:
        res.warn("🔴 有 %d 条**认不出是谁的**——请用 data.questions 弹窗请使用者选择或自定义主体"
                 % counts["needs"])
    else:
        res.warn("草稿好了。先处理 data.questions 的弹窗回答；随后再跑：\n"
                 "  office.py db-ingest --profile \"%s\"" % p_json)
    wr.log("db-absorb 读了 %s，产出 %d 条候选" % (src.name, len(rows)))
    return res
