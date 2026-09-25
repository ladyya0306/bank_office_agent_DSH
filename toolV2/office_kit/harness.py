"""管线：源文件 → 键值对 → 你复核 → 落库 → 匹配目标格 → 你勾选 → 保真填充。

这是**唯一一条**填表管线，跑在 **SCHEMA v2** 上（`27 §6`）。每一站都写库，
所以"上次你到底批了什么"不靠记忆，靠查库。

站点与人工干预点（**编号与 `27 §6` 一致**）

    ① 交付材料     → source（登记 + 存副本 + sha256）              `db-ingest`
    ② 解析提取     → fact（带出处、来源类型、置信度）              `db-ingest`
    ⓪ 可选识别资料 → 缺码或缺证件只提示，不阻止填报             `db-entity-key`
    ③ 数据规范化   → 去空格 / 单位归一（**只标记，不删除**）
    ④ 整理筛选     → 按键值对看库（置信度从高到低）                `db-show`
    ⑤ 用户复核     → 【干预点 1】源侧**四选**：接受/改后/丢弃/留空   `db-merge`
    ⑥ 写入数据库   → fact
    ⑦ 主体归一     → entity / entity_name                          `db-ingest`
    ⑧ 角色关卡     → case_role；源未明写 → 【干预点 2】🟡 提示级    `db-roles`
    ⑨ 匹配         → template_rule；<0.85 → 交模型判断 → 再问你     `db-propose`
    ⑩ 预演         → 【强制】dry-run：值/置信度/去向文档/去向格子   `db-fill --plan`
    ⑪ 校验         → 交付门禁 4 条 + 填报期 4 条 + 提示 7 条        `db-fill`
    ⑫ 填充         → 锚定替换（保真）；缺失留空                    `db-fill --apply-all/--select`
    ⑬ 签核         → 【干预点 3】使用者本人；一批一次 + 每份一次     `db-fill --sign`
    ⑭ 交付/导出    → 交付门禁（未签核等已实现的规则）

⚠️ **两条"拦"是分开的，不能混**（`27 §5.3` 的注）：

* **识别码和证件号码缺失 = 提示**，不阻止按工作区继续处理；
* **角色未明写 = 🟡 提示级**——只把**跟这个主体有关的那几个格子**留空并进 `review_queue`，
  **不阻断整批填报**。

⚠️ **本模块的诚实边界**：红线一的**完整**实现（逐值反查归属并与产物主体比对）还没做完。
本期落地的口径是 **角色前缀一致性**：`借款人*` 栏位的值必须来自扮演借款人的主体，
`保证人*` 同此；归属为空的值列为"**无法核对**"并计数，不假装核对过。
"""
from __future__ import annotations

import difflib
import json
import os
import re
import shutil
from pathlib import Path
from typing import Any

from . import doc_fill, doc_read, entity_key as EK, workroot as WR, xml_fill
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
from .store_v2 import (
    CASE_ROLES,
    NATURAL_PERSON_KEYS,
    StoreV2,
    SchemaError,
    allocate_batch,
    resolve_batch,
    sha256_file,
    valid_batch,
)

#: 置信度到了就自动采用（`27 §5.1` 的"自动填充率"来源）
AUTO_APPLY_THRESHOLD = 0.85
#: 低于这个连提都不用提给模型
PROPOSE_THRESHOLD = 0.45

#: 值就是主体的字段。**只有这些**字段需要问"这个主体的角色确认了没有"。
#: （早先的版本用一张 8 个字段的手写清单硬拦，把联系电话这种普通字段也一起挡了；
#:  正确的判据是"这个格子的**值本身就是一个主体**"。）
ROLE_DEPENDENT_KEYS = {
    "借款人名称", "保证人名称", "借款人法定代表人", "保证人法定代表人", "个人保证人姓名"}

#: 模板标签 → 字段键的同义词（`db-propose` 给候选打分用）
SYNONYMS: dict[str, list[str]] = {
    "借款人名称": ["客户名称", "企业名称", "借款企业名称", "客户企业名称", "申请人名称", "公司名称", "名称", "客户", "申请人"],
    "借款人住所": ["法定住所", "办公地点", "住所", "地址", "注册地址", "办公地址"],
    "借款人法定代表人": ["法定代表人", "法人", "法人代表", "法定代表人或授权代理人姓名"],
    "借款人法定代表人身份证": ["身份证号", "身份证号码"],
    "证件类型": ["身份证件类型", "有效证件类型"],
    "证件号码": ["证件号", "有效证件号码"],
    "联系电话": ["电话", "联系电话", "联系方式", "手机"],
    "联系人": ["联系人", "经办人", "业务联系人"],
    "联系人及电话": ["指定业务联系人姓名及电话", "联系人与电话"],
    "开户行": ["开户行", "开户银行", "开户行名称"],
    "账号": ["账号", "帐号", "银行账号", "账户"],
    "开户行及账号": ["在中信银行的开户行及账号", "开户行及账号", "开户银行及账号"],
    "流贷合同号": ["合同编号", "编号", "贷款合同号", "主合同编号"],
    "综合授信合同号": ["综合授信合同号", "授信合同号"],
    "业务编号": ["业务编号", "业务号"],
    "授信额度": ["授信额度", "额度", "授信金额", "额度金额"],
    "额度期限": ["额度期限", "授信期限", "期限", "额度起止日期"],
    "合同金额": ["合同金额", "金额"],
    "本次申请用信金额": ["用信金额", "申请用信金额", "用信业务币种及金额", "本次业务金额"],
    "用信用途": ["用途", "用信用途", "资金用途"],
    "保证人名称": ["保证人名称", "担保人名称"],
    "保证人住所": ["保证人住所", "担保人住所"],
    "授信额度/主合同编号": ["授信额度/主合同编号", "额度/主合同编号"],
    "已使用授信额度": ["已使用授信额度", "已用额度"],
    "可用授信额度": ["可使用授信额度", "可用授信额度", "可用额度"],
}


def _norm(s: str) -> str:
    return re.sub(r"[\s:：_\-（）()【】\[\]/]+", "", str(s)).lower()


def score_candidate(label: str, fact_key: str) -> float:
    """给"模板标签 ↔ 字段键"打分：1.0 完全相同 / 0.9 同义词 / 0.82 包含 / 0.5~0.8 模糊相似。

    ⚠️ **同义词先于"包含"判断**：`法定代表人` 是 `借款人法定代表人` 的同义词（0.9），
    而"是它的子串"只是巧合（0.82）。顺序反了会把真字段压到阈值以下，
    于是表单上明明有的格子被判成"没人认领"。
    """
    a, b = _norm(label), _norm(fact_key)
    if not a or not b:
        return 0.0
    if a == b:
        return 1.0
    for canonical, synonyms in SYNONYMS.items():
        na = {_norm(s) for s in synonyms}
        if b == _norm(canonical) and a in na:
            return 0.9
        if a == _norm(canonical) and b in na:
            return 0.9
    for synonyms in SYNONYMS.values():
        na = {_norm(s) for s in synonyms}
        if a in na and b in na:
            return 0.9
    if a in b or b in a:
        return 0.82
    return round(difflib.SequenceMatcher(None, a, b).ratio() * 0.8, 3)


def candidates_for(label: str, facts: dict[str, dict[str, Any]], *, top: int = 4) -> list[dict]:
    scored = []
    for key, meta in facts.items():
        s = score_candidate(label, key)
        if s >= PROPOSE_THRESHOLD:
            scored.append({"field": key, "score": s, "has_value": bool(meta.get("value"))})
    scored.sort(key=lambda d: -d["score"])
    return scored[:top]


def high_risk_class(field: str) -> str | None:
    """高风险五类（收款账号 / 金额 / 利率 / 日期 / 证件号码）。"""
    from .merge_v2 import high_risk_class as _hrc

    return _hrc(field)



def _open(args):
    """认领工作区 + 打开库。**工作区认不出来就把话说清楚，不猜。**"""
    try:
        wr = WR.open_for_command(args)
    except WR.WorkrootError as exc:
        raise OfficeKitError(str(exc)) from exc
    try:
        store = StoreV2(args._db, actor=getattr(args, "by", None)
                        or getattr(args, "actor", None) or "user")
    except Exception as exc:  # noqa: BLE001
        raise OfficeKitError(str(exc)) from exc
    wr.log("打开工作区 %s；库 %s" % (wr.path, args._db))
    return wr, store


def _rdir(args, command: str, batch_no: str | None = None) -> Path:
    """报告落哪：`<工作区>\\out\\<批次>\\_报告\\<命令>\\`（没工作区才退回老行为）。"""
    return WR.report_dir_for(args, command, batch_no=batch_no)


def _warn_outside(res: Result, args, *paths) -> None:
    """落点超出工作区就出**中文告警**（[24 §4.4 E]）。**只提醒，不拦**——文件夹不是边界。"""
    for w in WR.outside_warnings(args, *paths):
        res.warn(w)


def _tool_fp() -> str:
    """工具指纹（短）。**没有它，就说不清这份产物是哪个版本的工具做的。**"""
    try:
        from .fingerprint import short_fingerprint

        return short_fingerprint()
    except Exception:  # noqa: BLE001
        return "unknown"


# ==========================================================================
# 站点 ① ② ⑦ ⑧：源侧落库
# ==========================================================================
def source_text_index(paths: list[Path]) -> dict[str, str]:
    """源文档的逐行文本，键是 `第N行`。

    读**文档本身**是关键：profile 里的 `origin` 只是个指针（如 "0.1 第10行"）。
    早先的版本拿指针文本去搜角色词，于是永远搜不到，把每一份标得好好的材料
    都误报成"角色未明"。
    """
    index: dict[str, str] = {}
    for p in paths:
        if not p.exists():
            continue
        try:
            if p.suffix.lower() == ".docx":
                lines = [b.get("text", "") for b in doc_read.read_docx(p)["blocks"]]
            elif p.suffix.lower() in (".xlsx", ".xlsm", ".xls"):
                lines = []
                for block in doc_read.read_workbook(p)["sheets"]:
                    lines.append(" ".join(str(c) for c in block.get("header", [])))
                    for row in block.get("preview", []):
                        lines.append(" ".join(str(c) for c in row))
            else:
                continue
        except Exception:  # noqa: BLE001
            continue
        for i, text in enumerate(lines):
            if str(text).strip():
                index["第%d行" % i] = str(text)
                index["%s 第%d行" % (p.name, i)] = str(text)
    return index


#: 源文里紧挨在主体名左边的这些词，说明源文件**明写了**它的角色
ROLE_KEYWORDS: dict[str, list[str]] = {
    "借款人": ["借款人", "申请人", "客户", "用信人", "授信申请人"],
    "保证人": ["保证人", "担保人", "保证方"],
    "法定代表人": ["法定代表人", "法人代表", "法人"],
}


def infer_roles(store: StoreV2, batch_no: str, profile: dict[str, dict[str, Any]],
                source_index: dict[str, str] | None = None,
                *, source_id: int | None = None) -> dict[str, Any]:
    """源文件明写了角色 → 记 `case_role`；没写 → 🟡 生成一个问题（**不是硬拦**）。

    只有**主体**（人和公司）才有角色——金额、日期、合同号没有。
    角色只在源文里紧挨着主体名之前出现该角色词时才认定；否则问，
    因为"顾红军在文件里出现了两次"不是他是谁的法律代表人的证据。
    """
    source_index = source_index or {}
    entity_keys = ["借款人名称", "保证人名称", "借款人法定代表人", "保证人法定代表人",
                   "个人保证人姓名"]
    entities: list[str] = []
    for key in entity_keys:
        meta = profile.get(key)
        if meta and meta.get("value"):
            entities.append(str(meta["value"]))
    entities = list(dict.fromkeys(entities))

    by_line: dict[int, str] = {}
    for k, v in source_index.items():
        m = re.fullmatch(r"第(\d+)行", k)
        if m:
            by_line[int(m.group(1))] = v

    auto: list[dict[str, str]] = []
    for entity in entities:
        best: tuple[int, str, str] | None = None
        for _ln, line in sorted(by_line.items()):
            at = line.find(entity)
            if at < 0:
                continue
            prefix = line[:at]
            for role, words in ROLE_KEYWORDS.items():
                for w in words:
                    i = prefix.rfind(w)
                    if i < 0:
                        continue
                    if best is None or i > best[0]:
                        best = (i, role, line)
        if best is None:
            continue
        role, line = best[1], best[2]
        eid = store.find_entity_by_name(entity)
        if eid is None:
            continue
        store.set_role(eid, role, evidence=line.strip()[:120], decided_by="source",
                       source_id=source_id, batch_no=batch_no)
        auto.append({"entity": entity, "entity_id": eid, "role": role,
                     "evidence": line.strip()[:120]})

    known = {store.entity_label(r["entity_id"]) for r in store.roles()}
    questions: list[dict[str, Any]] = []
    for entity in entities:
        if entity in known:
            continue
        eid = store.find_entity_by_name(entity)
        qid = store.add_review(batch_no, "role", "role_unstated", entity_id=eid,
                               label=entity,
                               candidates=["借款人", "保证人", "法定代表人"],
                               field=None)
        questions.append({"id": qid, "entity": entity, "entity_id": eid})
    return {"entities": entities, "auto_roles": auto, "questions": questions,
            "gate_ok": not questions}


def cmd_db_ingest(args) -> Result:
    """① ② ⑦ ⑧：登记源文件、抽键值对、建主体、认角色——**一件事都不许漏记**。"""
    res = Result("db-ingest")
    wr, store = _open(args)

    profile_path = getattr(args, "profile", None)
    if not profile_path:
        store.close()
        raise OfficeKitError("--profile is required（要入库的键值对字典）")
    profile = doc_fill.load_profile(profile_path)

    batch_no = getattr(args, "batch", None)
    if batch_no and not valid_batch(batch_no):
        store.close()
        raise OfficeKitError("批次号必须是 YYYYMMDD-NN：%r" % batch_no)
    if not batch_no:
        batch_no = allocate_batch(store.conn)
    label = getattr(args, "label", None)

    copy_root = WR.copy_root_for(args)
    src_id, seen = store.register_source(Path(profile_path), batch_no, kind="profile",
                                         copy_root=copy_root, batch_label=label)
    if seen:
        res.warn("%s 之前已导入过（内容哈希相同），本次仍会刷新键值" % Path(profile_path).name)

    # 源文档：登记（可选存整份副本），并用来读角色
    source_paths: list[Path] = []
    for extra in getattr(args, "source", None) or []:
        for p in resolve_inputs(extra):
            source_paths.append(p)
            sid, was = store.register_source(
                p, batch_no, kind=p.suffix.lstrip("."),
                copy_root=copy_root, batch_label=label)
            res.data.setdefault("sources", []).append({"name": p.name, "already": was,
                                                       "source_id": sid})

    default_entity = getattr(args, "default_entity", None)

    # 第一遍：主体名字段先建主体，后面的字段才知道该挂给谁
    made_entities: list[dict] = []
    for key in ("借款人名称", "保证人名称", "借款人法定代表人", "保证人法定代表人",
                "个人保证人姓名"):
        meta = profile.get(key)
        if not meta or not meta.get("value"):
            continue
        eid, created = store.ensure_entity(str(meta["value"]),
                                           personal=key in NATURAL_PERSON_KEYS)
        if created:
            made_entities.append({"entity_id": eid, "name": str(meta["value"]),
                                  "from_key": key})

    # **角色槽位**：材料自己写明了"谁是借款人 / 保证人 / 法定代表人"，
    # 后面 `保证人住所`、`保证人联系电话` 这类键就照这几个槽位挂。
    #
    # ⚠️ 这里**绝不能**去查 `case_role` 表：那张表是下面 `infer_roles()` 才写的，
    #    而那时候事实已经写完了——查它必然拿到空，于是每一个"保证人*"的键
    #    都会掉进 default_entity，**挂到借款人头上**（实测：
    #    `保证人住所`/`保证人联系电话` 全挂给了借款人）。
    #    这是另一个会话在实操时抓出来的，已复现。
    role_slots: dict[str, int] = {}
    for role, keys in (("借款人", ("借款人名称",)),
                       ("保证人", ("保证人名称",)),
                       ("法定代表人", ("借款人法定代表人", "保证人法定代表人",
                                       "个人保证人姓名"))):
        for k in keys:
            meta = profile.get(k)
            if not meta or not meta.get("value"):
                continue
            eid = store.find_entity_by_name(str(meta["value"]))
            if eid is not None:
                role_slots[role] = eid
                break

    # 默认主体**必须在这一步之后**解析：它常常就是材料里刚建出来的那家
    default_eid: int | None = None
    if default_entity:
        default_eid = _resolve_entity(store, default_entity)

    added = missing = conflicts = unchanged = 0
    unassigned: list[str] = []
    for key, meta in profile.items():
        # 草稿里带的 `entity_name`（使用者用 `db-absorb --assign` 指明的）→ 解析成主体编号。
        # ⚠️ 这里**不能吞掉**：实测过用户指明了"这条是保证人的"，入库时被丢掉，
        #    结果还是按默认归给了借款人。
        declared_eid = None
        declared = str(meta.get("entity_name") or "").strip()
        if declared:
            declared_eid, created_d = store.ensure_entity(declared)
            if created_d:
                made_entities.append({"entity_id": declared_eid, "name": declared,
                                      "from_key": "%s（使用者指明）" % key})
        eid = _entity_for_fact(store, key, str(meta.get("value") or ""), default_eid,
                               role_slots, declared_eid)
        prov = meta.get("origin") or meta.get("source")
        if meta.get("missing"):
            out = store.put_fact(batch_no, key, None, entity_id=eid, source_id=src_id,
                                 source_kind="source", provenance=prov, status="missing",
                                 note=meta.get("note"), on_conflict="ignore")
            missing += 1
        else:
            # 使用者自己补的（`db-absorb --add`）必须记 `source_kind='user'`——红线二：
            # **不许把"你说的"伪装成"源文件里有的"**。
            kind_ = meta.get("source_kind") or "source"
            out = store.put_fact(batch_no, key, meta["value"], entity_id=eid, source_id=src_id,
                                 source_kind=kind_, provenance=prov, status="ok")
            if out["action"] == "conflict":
                conflicts += 1
            elif out["action"] == "unchanged":
                unchanged += 1
            else:
                added += 1
        if eid is None:
            unassigned.append(key)

    roles = infer_roles(store, batch_no, profile, source_text_index(source_paths),
                        source_id=src_id)

    # ⓪ 识别资料缺失仅提示，不影响后续填报。
    gate = EK.gate(store.conn)

    res.data.update({
        "batch_no": batch_no, "batch_label": label, "profile": Path(profile_path).name,
        "entities_created": made_entities,
        "facts_added": added, "facts_unchanged": unchanged, "facts_missing": missing,
        "conflicts": conflicts, "unassigned_keys": unassigned,
        "roles": roles,
        "identity_gate": gate,
        "summary": store.summary(),
    })
    if conflicts:
        res.warn("%d 个键库里已有不同的值——**没有覆盖**，已记成冲突。"
                 "要覆盖请跑 `db-merge` 出勾选清单" % conflicts)
    if unassigned:
        res.warn("%d 个键没能归属到具体主体（%s）——跨主体核对时它们会被列为"
                 "「无法核对」；可用 --default-entity 指明这批材料的默认主体"
                 % (len(unassigned), "、".join(unassigned[:5])))
    if roles["questions"]:
        res.warn("源文档未标明 %d 个主体的角色（🟡 提示级，不阻断）："
                 "相关格子会留空并进待确认清单，其余照填" % len(roles["questions"]))
    if gate["missing_count"]:
        res.warn("%d 个主体没有识别码或身份证号；这些资料可选，不影响继续处理。"
                 % gate["missing_count"])
    store.close()
    return res


def _resolve_entity(store: StoreV2, token: str) -> int:
    token = str(token).strip()
    if token.isdigit():
        eid = int(token)
        if store.conn.execute("SELECT id FROM entity WHERE id=?", (eid,)).fetchone() is None:
            raise OfficeKitError("没有这个主体：id=%s" % eid)
        return eid
    eid = store.find_entity_by_name(token)
    if eid is None:
        raise OfficeKitError("库里没有叫「%s」的主体——先用 db-entity-key --add 建档" % token)
    return eid


def _entity_for_fact(store: StoreV2, key: str, value: str, default_eid: int | None,
                     role_slots: dict[str, int] | None = None,
                     declared_eid: int | None = None) -> int | None:
    """这条键值对该算谁的。

    规则（**不猜**，而且**不依赖"角色已经认完了"**）：

    1. 主体名字段（`借款人名称` 等）→ 就是**它指的那个主体**；
    2. 键名声称了角色 → 先看材料里的**角色槽位**（`保证人名称` 点名的那家就是保证人，
       `借款人法定代表人` 点名的那个人就是法定代表人）；
       槽位没有才退回 `case_role` 表；
    3. 再没有 → `--default-entity` 指定的默认主体；
    4. 都没有 → 留空，并让报告把它算进"待归属"（`27 §7.1` 红线一的诚实口径）。

    ⚠️ **第 2 步的顺序是有血的教训的**：早先只看 `case_role`，而那张表是**写事实之后**
    才填的，于是所有 `保证人住所` / `保证人联系电话` 这类键**全部挂到了借款人头上**。
    键名声称的角色，从**材料自己写的名字**就能定，不该等角色推断。
    """
    if key in ROLE_DEPENDENT_KEYS and value:
        eid = store.find_entity_by_name(value)
        if eid is not None:
            return eid
    # **使用者自己指明的归属**（`db-absorb --assign`）：他说这条是谁的，就是谁的。
    # 优先级排在"主体名键"之后——主体名键的值本身就定义了主体，没有歧义；
    # 其余情况**使用者的话最大**。
    if declared_eid is not None:
        return declared_eid
    claimed = _role_claimed_by_field(key)
    if claimed:
        slots = role_slots or {}
        if claimed in slots:
            return slots[claimed]
        eid = store.entity_for_role(claimed)
        if eid is not None:
            return eid
    return default_eid


# ==========================================================================
# 站点 ⑧：角色关卡（🟡 提示级）
# ==========================================================================
def cmd_db_roles(args) -> Result:
    """看、或回答"谁是借款人 / 保证人 / 法定代表人"。

    ⚠️ 这是 **🟡 提示级**（`26 §2.12`）：没确认只让**相关的那几个格子**留空，
    **不阻断整批填报**。识别资料缺失也不阻断。
    """
    res = Result("db-roles")
    wr, store = _open(args)
    batch_no = getattr(args, "batch", None) or store.current_batch() or allocate_batch(store.conn)

    rid = getattr(args, "id", None)
    if rid:
        answer = getattr(args, "answer", None)
        if not answer:
            store.close()
            raise OfficeKitError("给出了 --id 就要给 --answer")
        if answer not in CASE_ROLES:
            # 也接受旧写法「借款人的法定代表人」，映射到粗粒度三选一
            mapped = ("法定代表人" if "法定代表人" in answer else
                      "保证人" if "保证人" in answer else
                      "借款人" if "借款人" in answer else None)
            if mapped is None:
                store.close()
                raise OfficeKitError("角色只能是 %s（X-3 粗粒度）：%r"
                                     % ("/".join(CASE_ROLES), answer))
            answer = mapped
        row = store.conn.execute("SELECT * FROM review_queue WHERE id=?", (int(rid),)).fetchone()
        if row is None:
            store.close()
            raise OfficeKitError("没有这个待确认项：id=%s" % rid)
        eid = row["entity_id"] or (store.find_entity_by_name(row["label"] or "") if row["label"] else None)
        if eid is None:
            store.close()
            raise OfficeKitError("这一项没关联到主体，没法记角色（id=%s）" % rid)
        store.resolve_review(int(rid), answer, answered_value=answer,
                            decided_by=getattr(args, "by", None) or "user")
        store.set_role(int(eid), answer, evidence="人工确认（来自待确认项 #%s）" % rid,
                       decided_by="human", batch_no=batch_no)
        res.data.update({"resolved": int(rid), "answer": answer, "entity_id": int(eid)})
    elif getattr(args, "entity", None) and getattr(args, "role", None):
        eid = _resolve_entity(store, args.entity)
        store.set_role(eid, args.role, evidence="人工指定", decided_by="human",
                       batch_no=batch_no)
        res.data.update({"set": {"entity_id": eid, "role": args.role}})

    res.data["batch_no"] = batch_no
    res.data["roles"] = store.roles()
    res.data["open_questions"] = [r for r in store.open_reviews(kind="role")]
    res.data["gate_ok"] = not res.data["open_questions"]
    res.data["summary"] = store.summary()
    if not res.data["gate_ok"]:
        res.warn("还有 %d 个主体的角色没确认（🟡 提示级，不阻断填报）："
                 "跟这些主体有关的格子会留空" % len(res.data["open_questions"]))
    store.close()
    return res


# ==========================================================================
# 站点 ⑨：匹配（按模板标签生成候选规则）
# ==========================================================================
LABEL_RE = re.compile(r"([\u4e00-\u9fffA-Za-z*＊（）()、·]{2,24})\s*[:：]")


def extract_labels(path: Path, known_labels: set[str] | None = None) -> list[dict[str, Any]]:
    """在模板里找**可能是待填格子**的标签。

    识别冒号标签，以及表格中与已知字段一致的无冒号标签。是否真的可填，留给 :func:`score_candidate` 判——
    抽取阶段故意宽松，因为在这里过滤会把真字段悄悄藏起来
    （早先的版本只留"冒号后面是空白"的，于是所有已经预填了值的表单字段全被漏掉）。

    **表格里的标签也要认**：银行表单绝大多数是"左边一格是标签、右边一格留空"，
    只扫段落会把整张表漏掉。表格标签的落点分两种：

    * 右边那一格是空的 → 就填那一格（`kind="cell"`，最贴近表单原意）；
    * 右边那格有东西 → 在标签所在的段落里**接着标签写**（`kind="anchor"`），
      这样标签本身不会被覆盖掉。
    """
    known = {_norm(str(x)) for x in (known_labels or set()) if str(x).strip()}

    if path.suffix.lower() == ".xlsx":
        import openpyxl

        wb = openpyxl.load_workbook(path)
        labels: list[dict[str, Any]] = []
        try:
            for ws in wb.worksheets:
                # Sparse sheets may report enormous rectangular dimensions.
                for cell in list(ws._cells.values()):
                    if not isinstance(cell.value, str):
                        continue
                    text = cell.value
                    matches = list(LABEL_RE.finditer(text))
                    # 无冒号时只接受数据库中的已知字段名，避免把普通正文当成标签。
                    if not matches and _norm(text) in known:
                        matches = [None]
                    for m in matches:
                        if m is None:
                            label = text.strip().lstrip("*＊").strip()
                            required = text.strip().startswith(("*", "＊"))
                            anchor = text.strip()
                        else:
                            label = m.group(1).replace("*", "").replace("＊", "").strip()
                            required = "*" in m.group(1) or "＊" in m.group(1)
                            anchor = m.group(0).strip()
                        label = re.sub(r"^(?:(?:[一二三四五六七八九十]+|[0-9]+)[、.．]|、)\s*", "", label)
                        if not label:
                            continue
                        target = {"kind": "xlsx_cell", "sheet": ws.title,
                                  "cell": cell.coordinate, "anchor": anchor}
                        # 无冒号标签使用右侧空单元格；合并标签从合并区域之后开始，
                        # 绝不把值写回标签自身或同一合并区域。
                        if m is None:
                            merged_label = next((r for r in ws.merged_cells.ranges
                                                 if cell.coordinate in r), None)
                            start_col = (merged_label.max_col + 1
                                         if merged_label else cell.column + 1)
                            if start_col <= ws.max_column:
                                candidate = ws.cell(cell.row, start_col)
                                candidate_has_format = bool(candidate.has_style)
                                if candidate.value in (None, "") and not candidate.data_type == "f" and candidate_has_format:
                                    target["cell"] = candidate.coordinate
                                    target["anchor"] = None
                                    target["label_cell"] = cell.coordinate
                                else:
                                    continue
                            else:
                                continue
                        labels.append({"text": label,
                                       "location": f"{ws.title}!{cell.coordinate}",
                                       "is_required": int(required), "target": target})
        finally:
            wb.close()
        return labels

    doc = doc_read.read_docx(path)
    found: list[dict[str, Any]] = []
    seen: set[tuple] = set()

    def add(raw_label: str, location: str, target: dict) -> None:
        required = ("*" in raw_label) or ("＊" in raw_label)
        label = raw_label.replace("*", "").replace("＊", "").strip()
        # 明确的列表前缀属于版式编号，不是字段名；anchor 仍保留原文。
        label = re.sub(r"^(?:(?:[一二三四五六七八九十]+|[0-9]+)[、.．]|、)\s*", "", label)
        if not label:
            return
        identity = (location, target.get("part", ""), target.get("kind", ""),
                    target.get("table"), target.get("row"), target.get("col"),
                    target.get("anchor"))
        if identity in seen:
            return
        seen.add(identity)
        found.append({"text": label, "location": location,
                      "is_required": 1 if required else 0, "target": target})

    for bi, block in enumerate(doc["blocks"]):
        if block["type"] not in ("paragraph", "list_item"):
            continue
        text = block.get("text", "")
        if not text:
            continue
        for m in LABEL_RE.finditer(text):
            if len(text[m.end():].strip()) > 30:
                continue  # 冒号后面是一整句话 → 那是正文，不是格子
            add(m.group(1), "block[%d]" % bi,
                {"kind": "anchor", "anchor": m.group(0).strip(), "max_blank": 200})

    if path.suffix.lower() == ".docx":
        import docx as _docx

        d = _docx.Document(str(path))
        for ti, tbl in enumerate(d.tables):
            for ri, row in enumerate(tbl.rows):
                cells = row.cells
                seen_cells: set[int] = set()
                for ci, cell in enumerate(cells):
                    tc_identity = id(cell._tc)
                    if tc_identity in seen_cells:
                        continue
                    seen_cells.add(tc_identity)
                    cell_text = cell.text.strip()
                    matches = list(LABEL_RE.finditer(cell_text))
                    if not matches and _norm(cell_text) in known:
                        matches = [None]
                    for m in matches:
                        if m is None:
                            label = cell_text.lstrip("*＊").strip()
                            required = cell_text.startswith(("*", "＊"))
                            anchor = cell_text
                        else:
                            label = m.group(1)
                            required = "*" in m.group(1) or "＊" in m.group(1)
                            anchor = m.group(0).strip()
                        if not label:
                            continue
                        if m is not None and len(cell_text[m.end():].strip()) > 30:
                            continue
                        # python-docx 会为横向合并单元格返回多个 Cell 对象；
                        # 用底层 w:tc 身份判断，不能用 Python 对象 is。
                        nxt = None
                        for ni in range(ci + 1, len(cells)):
                            if cells[ni]._tc is not cell._tc:
                                nxt = cells[ni]
                                break
                        if nxt is not None and not nxt.text.strip() and nxt._tc is not cell._tc:
                            next_col = next(ni for ni in range(ci + 1, len(cells))
                                            if cells[ni]._tc is nxt._tc)
                            add(label, "table[%d].cell(%d,%d)" % (ti, ri, ci),
                                {"kind": "cell", "table": ti, "row": ri,
                                 "col": next_col})
                        elif m is not None:
                            add(label, "table[%d].cell(%d,%d)" % (ti, ri, ci),
                                {"kind": "anchor", "anchor": anchor,
                                 "max_blank": 200})
        # 页眉/页脚只扫描压缩包里已经存在的 XML 部件，避免访问 sections
        # 造成空页眉物化。XmlEngine 后续会按 part 精确查找与写入。
        import zipfile
        from xml.etree import ElementTree as ET
        with zipfile.ZipFile(str(path)) as zf:
            part_names = sorted(n for n in zf.namelist()
                                if re.match(r"word/(?:header|footer)\d+\.xml$", n))
            for part in part_names:
                root = ET.fromstring(zf.read(part))
                ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
                for pi, paragraph in enumerate(root.findall(".//w:p", ns)):
                    text = "".join((node.text or "") for node in paragraph.findall(".//w:t", ns))
                    if not text:
                        continue
                    for m in LABEL_RE.finditer(text):
                        if len(text[m.end():].strip()) > 30:
                            continue
                        add(m.group(1), f"{part}[{pi}]",
                            {"kind": "anchor", "anchor": m.group(0).strip(),
                             "part": part, "max_blank": 200})
                    if _norm(text) in known and not LABEL_RE.search(text):
                        add(text, f"{part}[{pi}]",
                            {"kind": "anchor", "anchor": text.strip(),
                             "part": part, "max_blank": 200})
    return found


def cmd_db_propose(args) -> Result:
    """给每份模板的每个标签找最合适的字段，并按置信度分级。"""
    res = Result("db-propose")
    wr, store = _open(args)
    batch_no = getattr(args, "batch", None) or store.current_batch() or allocate_batch(store.conn)
    store.batch_scope = batch_no
    facts = store.facts_by_key()
    if not facts:
        store.close()
        raise OfficeKitError("数据库里还没有键值对，先跑 db-ingest")

    proposals: list[dict[str, Any]] = []
    auto = 0
    existing_kept = 0
    reused_templates = 0
    for tpl_path in resolve_inputs(args.input):
        from .rule_pack import reuse_verified, suggest_similar
        reuse = reuse_verified(store, tpl_path, batch_no)
        if reuse and not reuse["conflicts_skipped"]:
            reused_templates += 1
            continue
        if reuse and reuse["conflicts_skipped"]:
            res.warn("%s：本工作区已有 %d 条不同的填写位置，已保留本工作区规则，请核对。"
                     % (tpl_path.name, reuse["conflicts_skipped"]))
        if not reuse:
            similar = suggest_similar(tpl_path)
            if similar:
                res.warn("%s：已复核目录里有文件名相近的旧版（%s），仅供参考；内容不同，未自动套用。"
                         % (tpl_path.name, "、".join(similar)))
        tid = store.register_template(tpl_path, batch_no)
        original_rules = {r['field']: r for r in store.rules_for(tpl_path)}
        existing_fields = set(original_rules)
        chosen_locations: dict[str, list[str]] = {}
        known_labels = set(facts)
        for aliases in SYNONYMS.values():
            known_labels.update(aliases)
        for lab in extract_labels(tpl_path, known_labels=known_labels):
            cands = candidates_for(lab["text"], facts)
            best = cands[0] if cands else None
            entry = {
                "template": tpl_path.name, "label": lab["text"],
                "location": lab["location"], "target": lab["target"],
                "is_required": lab["is_required"],
                "candidates": cands,
                "chosen": best["field"] if best and best["score"] >= AUTO_APPLY_THRESHOLD else None,
                "confidence": best["score"] if best else 0.0,
            }
            if entry["chosen"]:
                locations = chosen_locations.setdefault(entry["chosen"], [])
                locations.append(entry["location"])
                if entry["chosen"] in existing_fields:
                    existing_kept += 1
                else:
                    store.add_rule(tid, entry["chosen"], lab["text"], lab["target"],
                                   is_required=lab["is_required"], match_kind="exact",
                                   confidence=entry["confidence"], decided_by="auto",
                                   batch_no=batch_no)
                    existing_fields.add(entry["chosen"])
                    auto += 1
            proposals.append(entry)

        # One field can have several explicit locations (e.g. two worksheets).
        # Keep one DB rule, enumerate its targets in the plan, and ask once with
        # all locations listed. Never discard/replace a manually chosen target.
        for field in chosen_locations:
            items = [p for p in proposals if p['template'] == tpl_path.name and p['chosen'] == field]
            targets = list({json.dumps(p['target'], sort_keys=True): p['target'] for p in items}.values())
            if len(targets) < 2:
                continue
            old = original_rules.get(field)
            if old and old.get('decided_by') != 'auto':
                res.warn(f'{tpl_path.name}：{field} 有其他候选位置，保留原复核规则，请核对未覆盖位置。')
                continue
            if old:
                previous = old['target'].get('targets', [old['target']])
                # Additional automatic locations must not replace an existing
                # rule whose meaning no longer agrees with the current template.
                if any(t not in targets for t in previous):
                    continue
            store.add_rule(tid, field, items[0]['label'], {'kind': 'multi', 'targets': targets},
                           is_required=max(p['is_required'] for p in items),
                           confidence=min(p['confidence'] for p in items),
                           decided_by='auto', batch_no=batch_no)

    out = _rdir(args, "db-propose", batch_no)
    p = out / "proposals.json"
    write_text(p, json.dumps(proposals, ensure_ascii=False, indent=2))
    res.add_artifact(p, "候选规则（含置信度，供大模型/人工裁决）")

    needs_model = [x for x in proposals if not x["chosen"] and x["candidates"]]
    no_candidate = [x for x in proposals if not x["candidates"]]
    res.data.update({"batch_no": batch_no, "labels": len(proposals), "auto_rules": auto,
                     "existing_rules_kept": existing_kept,
                     "reused_verified_templates": reused_templates,
                     "needs_model_arbitration": len(needs_model),
                     "no_candidate": len(no_candidate), "proposals": proposals,
                     "summary": store.summary()})
    if needs_model:
        res.warn("%d 个标签置信度不足，需要大模型裁决；请审阅 proposals.json 的 candidates "
                 "后用 db-rule 写入决定" % len(needs_model))
    store.close()
    return res


def cmd_db_rule(args) -> Result:
    """写入一条人工/模型裁决的规则（一行 = 目标表单上的一个格子）。"""
    res = Result("db-rule")
    wr, store = _open(args)
    batch_no = getattr(args, "batch", None) or store.current_batch() or allocate_batch(store.conn)
    tpl = Path(args.template)
    if not tpl.exists():
        store.close()
        raise OfficeKitError("模板不存在：%s" % tpl)
    tid = store.register_template(tpl, batch_no)
    target = json.loads(args.target) if getattr(args, "target", None) else {
        "kind": "anchor", "anchor": args.anchor, "max_blank": 200,
    }
    if getattr(args, "before", None):
        target["before"] = args.before
    rid = store.add_rule(tid, args.field, getattr(args, "label", None) or args.anchor, target,
                         is_required=1 if getattr(args, "required", False) else 0,
                         match_kind=getattr(args, "match_kind", None) or "exact",
                         confidence=getattr(args, "confidence", None),
                         decided_by=getattr(args, "decided_by", None) or "human",
                         batch_no=batch_no)
    res.data.update({"rule_id": rid, "template": tpl.name, "field": args.field,
                     "target": target, "summary": store.summary()})
    store.close()
    return res


def cmd_db_rule_disable(args) -> Result:
    """Retire a mistaken position without deleting its history or using row IDs."""
    res = Result("db-rule-disable")
    _wr, store = _open(args)
    try:
        try:
            expected = json.loads(args.expected_target)
        except ValueError as exc:
            raise OfficeKitError("--expected-target 必须是当前位置的 JSON") from exc
        if not isinstance(expected, dict) or not args.reason.strip():
            raise OfficeKitError("请提供当前位置和停用原因")
        try:
            rid = store.disable_rule(args.template, args.field, expected, args.reason.strip())
        except SchemaError as exc:
            raise OfficeKitError(str(exc)) from exc
        res.data.update({"rule_id_retained": rid, "template": Path(args.template).name,
                         "field": args.field, "reason": args.reason.strip()})
        return res
    finally:
        store.close()


# ==========================================================================
# 站点 ⑩ ⑪ ⑫ ⑬：预演 → 校验 → 填充 → 签核
# ==========================================================================
def _gate_identity(store: StoreV2) -> dict:
    """⓪ 可选识别资料缺失提示。"""
    return EK.gate(store.conn)


def _template_structure(tpl: Path) -> tuple[int, str, str]:
    """模板结构自检：只支持单一表头 / 无公式 / 无宏 / 无外链（`23 §3.1`）。"""
    kind = tpl.suffix.lower().lstrip(".")
    if kind not in ("docx", "xlsx"):
        return 0, kind, "本期只支持 .docx 与 .xlsx 作目标模板（旧版 .doc/.xls 请先 to-docx）"
    try:
        parts = xml_fill.package_parts(tpl)
    except Exception as exc:  # noqa: BLE001
        return 0, kind, "打不开：%s" % exc
    notes: list[str] = []
    if any("vbaProject" in n for n in parts):
        notes.append("含宏（vbaProject）")
    if any(n.endswith(".rels") and b'TargetMode="External"' in parts[n] for n in parts):
        notes.append("含外部链接")
    if kind == "xlsx" and any(
            (b"<f>" in parts[n] or b"<f " in parts[n])
            for n in parts if n.startswith("xl/worksheets/")):
        notes.append("含公式")
    if notes:
        return 0, kind, "；".join(notes)
    return 1, kind, "结构正常"


def _base_facts(store: StoreV2) -> dict[str, dict]:
    return store.facts_by_key()


def _open_questions_by_entity(store: StoreV2) -> set[int | None]:
    return {r["entity_id"] for r in store.open_reviews(kind="role")}


def _tied_candidates(label: str, facts: dict[str, dict], chosen: str) -> list[str]:
    """并列最高分的字段（不含已选中的那个）。

    `27 §5.1`：**分数并列 → 直接问你**。表单上只写「名称：」时，
    `借款人名称` 与 `保证人名称` 都是 0.9——这时候替你挑一个就是猜。
    """
    scored = sorted(((score_candidate(label, k), k) for k in facts), reverse=True)
    if not scored:
        return []
    top = scored[0][0]
    if top < PROPOSE_THRESHOLD:
        return []
    return [k for s, k in scored if abs(s - top) < 1e-9 and k != chosen]


def _template_subject_scope(store: StoreV2, template: Path,
                            rules: list[dict]) -> tuple[int | None, list[int] | None, str, str]:
    """Determine a template's subject without choosing the first guarantor.

    A role written in the filename is the strongest template declaration.  A
    single role claimed by its rules is only a fallback.  If either declaration
    names more than one role, the template has no unique subject declaration.
    """
    name = template.stem
    filename_roles = {role for role in ("借款人", "保证人") if role in name}
    rule_roles = {_role_claimed_by_field(r["field"]) for r in rules}
    rule_roles.discard(None)

    role: str | None = None
    declaration = ""
    if len(filename_roles) == 1:
        role = next(iter(filename_roles))
        declaration = "模板文件名明确为%s材料" % role
    elif len(filename_roles) > 1:
        return None, None, "认不出", "模板文件名同时出现借款人和保证人，不能优先选择其中一个"
    elif len(rule_roles) == 1:
        candidate = next(iter(rule_roles))
        if candidate in ("借款人", "保证人"):
            role = candidate
            declaration = "模板规则只明确声明%s" % role
        else:
            return None, None, "认不出", "模板规则只声明法定代表人，未说明关联的借款人或保证人"
    elif len(rule_roles) > 1:
        return None, None, "认不出", "模板规则声明多个主体角色，不能优先选择借款人"
    else:
        return None, None, "认不出", "模板文件名和规则均未声明借款人或保证人"

    # A legal-representative certificate or board resolution is a company
    # document. Use explicit source evidence, not legacy default entity types.
    company_material = any(token in name for token in ("法定代表人身份证明", "董事会决议"))
    candidates = store.entities_for_role(role)
    if company_material:
        companies = [eid for eid in candidates if store.is_company_evidenced(eid)]
        if not companies and candidates:
            # Lack of company evidence is not lack of an address/phone. Keep
            # candidates for explicit selection rather than hiding their facts.
            candidates = [eid for eid in candidates if store.conn.execute(
                'SELECT entity_type FROM entity WHERE id=?', (eid,)).fetchone()[0] != 'natural_person']
            return None, candidates, role, declaration + '；源材料未明确企业身份，请在本表问题中选择适用主体'
        candidates = companies
    type_note = ("；该模板是明确公司材料，已仅保留有企业来源证据的主体"
                 if company_material else "")
    if len(candidates) == 1:
        return candidates[0], candidates, role, declaration + type_note
    if len(candidates) > 1:
        return None, candidates, role, declaration + type_note + "；有%d个候选主体，需逐模板确认" % len(candidates)
    return None, [], role, declaration + type_note + "；库中没有符合条件的主体"


def _template_subject_eid(store: StoreV2, rules: list[dict],
                          template: Path | None = None) -> tuple[int | None, str]:
    """Compatibility wrapper for callers that only need the unique subject."""
    if template is None:
        return None, "认不出"
    eid, _scope, label, _note = _template_subject_scope(store, template, rules)
    return eid, label


def build_fill_plan(store: StoreV2, templates: list[Path], *, batch_no: str,
                    run_id: str) -> dict:
    """⑩ **强制预演**：每个格子填什么、几分把握、为什么要问你——一屏看完。

    `decision`：
      * `auto` —— 置信度 ≥ 0.85、不是高风险五类、角色已确认、有值 → `--apply-all` 会填
      * `ask`  —— 必须**逐个勾**（`ask_reason` 说明为什么）
      * `empty`——源文件里确实没有这个值 → 留空 + 进待确认清单
    """
    store.batch_scope = batch_no
    open_roles = _open_questions_by_entity(store)
    rows: list[dict] = []
    for tpl in templates:
        template_hash = sha256_file(tpl)
        tid = store.register_template(tpl, batch_no)
        struct_ok, kind, struct_note = _template_structure(tpl)
        rules = []
        for saved in store.rules_for(str(tpl)):
            target = saved['target']
            if target.get('kind') == 'multi':
                targets = target.get('targets')
                if not isinstance(targets, list) or not targets or any(
                        not isinstance(t, dict) or t.get('kind') not in ('cell', 'anchor', 'xlsx_cell')
                        for t in targets):
                    raise OfficeKitError('多位置规则必须列出明确的单层填写位置')
                rules.extend({**saved, 'target': t} for t in targets)
            else:
                rules.append(saved)
        if not rules:
            rows.append({"template": tpl.name, "template_id": tid, "kind": "template_note",
                         "template_sha256": template_hash,
                         "note": "数据库里没有规则（先跑 db-propose 或 db-rule）",
                         "structure_ok": struct_ok, "structure_note": struct_note})
            continue
        # ★ **按本次主体取值**：同一键属于多个主体时，这里要么挑对，要么留成"问你"
        subject_eid, subject_scope, subject_label, subject_note = _template_subject_scope(
            store, tpl, rules)
        facts = store.facts_for_subject(subject_eid, scope_entity_ids=subject_scope)
        from .fact_catalog import qualified_facts
        facts.update({key: meta for key, meta in qualified_facts(store).items()
                      if meta.get('_qualified')})
        for rule in rules:
            field = rule["field"]
            meta = facts.get(field)
            value = None if meta is None else meta.get("value")
            conf = rule.get("confidence")
            if conf is None:
                conf = 1.0 if meta is not None else 0.0
            hr = high_risk_class(field)
            eid = None if meta is None else meta.get("entity_id")
            role_blocked = field in ROLE_DEPENDENT_KEYS and eid in open_roles
            tied = _tied_candidates(rule.get("label") or field, facts, field)
            ambiguous = bool(meta is not None and meta.get("_ambiguous"))
            cands = (meta or {}).get("_candidates") or []
            if ambiguous:
                # 🔴 库里这个键**属于多个主体**，而字段名又没指明是谁的 —— **不许替你挑**
                who = "、".join("%s（%s）" % (c["entity_name"], c["value"]) for c in cands[:3])
                decision, ask = "ask", (
                    "库里这个键有 %d 个主体的值：%s。字段名没指明是谁的 —— "
                    "请用 `--use <本表编号>=<主体名或#编号>` 指明" % (len(cands), who))
            elif value is None or not str(value).strip():
                decision, ask = "empty", "源文件里没有这个值（缺失留空）"
            elif role_blocked:
                decision, ask = "ask", "这个主体的角色还没确认（🟡 提示级）"
            elif tied:
                # `27 §5.1`：分数并列 → 直接问你，不许替你挑一个
                decision, ask = "ask", "候选得分并列（%s 都是 %.2f），请你指明是哪一格" % (
                    "、".join(tied), conf)
            elif conf < AUTO_APPLY_THRESHOLD:
                decision, ask = "ask", "匹配置信度 %.2f < %.2f" % (conf, AUTO_APPLY_THRESHOLD)
            elif hr:
                decision, ask = "ask", "高风险五类字段（%s），无论多高分都要你点头" % hr
            else:
                decision, ask = "auto", ""
            rows.append({
                "n": 0, "template": tpl.name, "template_id": tid, "template_rule_id": rule["id"],
                "template_sha256": template_hash,
                "field": field, "label": rule.get("label") or field,
                "target": rule["target"], "value": None if value is None else str(value),
                "confidence": conf, "is_required": bool(rule.get("is_required")),
                "high_risk": hr, "entity_id": eid,
                "entity_name": ("（多个主体，待你指明）" if ambiguous
                                else store.entity_label(eid)),
                "provenance": None if meta is None else meta.get("provenance"),
                "source_kind": None if meta is None else meta.get("source_kind"),
                "decision": decision, "ask_reason": ask,
                "ambiguous": ambiguous, "candidates": cands,
                "subject_eid": subject_eid, "subject_label": subject_label,
                "subject_scope": subject_scope, "subject_note": subject_note,
                "qualified_source": bool((meta or {}).get('_qualified')),
                "relation_owner_eid": (meta or {}).get('_relation_owner_eid'),
                "structure_ok": struct_ok, "structure_note": struct_note, "doc_kind": kind,
                "already_filled": _already_filled(store, field, batch_no),
                "kind": "slot",
            })
    n = 0
    for row in rows:
        if row["kind"] == "slot":
            n += 1
            row["n"] = n
    filled = [r for r in rows if r["kind"] == "slot" and r["decision"] != "empty"]
    return {
        "stage": "plan", "batch_no": batch_no, "run_id": run_id,
        "rows": rows,
        "counts": {
            "templates": len(templates),
            "slots": len(rows),
            "auto": sum(1 for r in rows if r.get("decision") == "auto"),
            "ask": sum(1 for r in rows if r.get("decision") == "ask"),
            "empty": sum(1 for r in rows if r.get("decision") == "empty"),
            "candidates": len(filled),
        },
    }


def _already_filled(store: StoreV2, field: str, batch_no: str) -> list[str]:
    """这个东西以前填进过哪些产物（只做提示，不阻断）。"""
    rows = store.conn.execute(
        "SELECT DISTINCT a.artifact_path FROM fill_op f JOIN fill_op a"
        " ON a.run_id=f.run_id AND a.kind='artifact'"
        " WHERE f.kind='field' AND f.field=? AND f.status='ok'", (field,)).fetchall()
    return sorted({r[0] for r in rows if r[0]})


def uncovered_xlsx_labels(plan: dict, proposals: list[dict]) -> list[dict]:
    """Warn when a clear Excel label lost its rule during later edits.

    A proposal points at a label cell; the actual value may belong in its
    neighboring cell. Compare the field or visible label, not the coordinate.
    """
    rows = [r for r in plan["rows"] if r.get("kind") == "slot"]
    seen = set()
    missing = []
    for item in proposals:
        if not isinstance(item, dict):
            continue
        tpl, label, field = item.get("template"), item.get("label"), item.get("chosen")
        if not str(tpl).lower().endswith(".xlsx") or not label or not field:
            continue
        if float(item.get("confidence") or 0) < AUTO_APPLY_THRESHOLD:
            continue
        key = (tpl, _norm(label))
        if key in seen:
            continue
        seen.add(key)
        if any(r["template"] == tpl and (r["field"] == field or
               (_norm(label) in _norm(r.get("label") or "") and len(_norm(label)) >= 2))
               for r in rows):
            continue
        missing.append({"template": tpl, "label": label,
                        "suggested_field": field, "location": item.get("location")})
    return missing


def current_xlsx_gaps(store: StoreV2, templates: list[Path], plan: dict) -> list[dict]:
    """Recheck clear Excel labels against current rules, without an old proposal file."""
    facts = store.facts_by_key()
    proposals = []
    for tpl in templates:
        if tpl.suffix.lower() != ".xlsx":
            continue
        if plan.get('position_inventory', {}).get('template_hash') == sha256_file(tpl):
            # The workflow has accounted for individual locations, including
            # intentional blanks and differently qualified source-field names.
            continue
        known_labels = set(facts)
        for aliases in SYNONYMS.values():
            known_labels.update(aliases)
        for label in extract_labels(tpl, known_labels=known_labels):
            candidates = candidates_for(label["text"], facts)
            if not candidates or candidates[0]["score"] < AUTO_APPLY_THRESHOLD:
                continue
            # Equal top scores are ambiguous: do not guess which field belongs here.
            if len(candidates) > 1 and candidates[0]["score"] == candidates[1]["score"]:
                continue
            proposals.append({"template": tpl.name, "label": label["text"],
                              "location": label["location"],
                              "target": label["target"],
                              "chosen": candidates[0]["field"],
                              "confidence": candidates[0]["score"]})
    gaps = []
    for proposal in proposals:
        target = proposal.get("target") or {}
        rows = [r for r in plan.get("rows", [])
                if r.get("kind") == "slot" and r.get("template") == proposal["template"]
                and r.get("field") == proposal["chosen"]]
        matched = False
        for row in rows:
            actual = row.get("target") or {}
            if (actual.get("kind") == "xlsx_cell" and
                    actual.get("sheet") == target.get("sheet") and
                    actual.get("cell") == target.get("cell") and
                    (actual.get("anchor") or "") == (target.get("anchor") or "")):
                matched = True
                break
        if not matched:
            gaps.append({"template": proposal["template"], "label": proposal["label"],
                         "suggested_field": proposal["chosen"],
                         "location": proposal["location"], "target": target})
    return gaps


def current_docx_target_errors(templates: list[Path], plan: dict) -> list[dict]:
    """Check Word rule locations before asking the user to approve a fill."""
    errors = []
    for tpl in templates:
        if tpl.suffix.lower() != ".docx":
            continue
        rows = [r for r in plan["rows"] if r["kind"] == "slot" and r["template"] == tpl.name]
        if not rows:
            continue
        try:
            engine = doc_fill.XmlEngine(tpl)
        except Exception as exc:  # noqa: BLE001
            errors.append({"template": tpl.name, "field": "", "reason": "无法检查 Word 模板：%s" % exc})
            continue
        def positions(target):
            if target.get("kind") == "cell":
                cell = engine.document.tables[int(target["table"])].cell(
                    int(target["row"]), int(target["col"]))
                return {(cell.paragraphs[0]._element, 0)}
            return {(p._element, start) for p, start, _ in engine.find_anchor_hits(target)}

        covered = {}
        for row in rows:
            target = row["target"]
            try:
                if target.get("kind") == "cell":
                    cell = engine.document.tables[int(target["table"])].cell(
                        int(target["row"]), int(target["col"]))
                    if not cell.paragraphs or engine.xml_for(cell.paragraphs[0]) is None:
                        raise OfficeKitError("目标单元格无法对应到原文件")
                else:
                    hits = engine.find_anchor_hits(target)
                    if not hits:
                        raise OfficeKitError("找不到填写位置")
                    if len(hits) > int(target.get("max_matches", 1)) and not target.get("only_first"):
                        raise OfficeKitError("找到 %d 个位置，无法确定填哪一处" % len(hits))
                covered.setdefault(row['field'], set()).update(positions(target))
            except (OfficeKitError, IndexError, KeyError, ValueError, TypeError) as exc:
                errors.append({"template": tpl.name, "field": row["field"], "reason": str(exc)})
        if plan.get('position_inventory', {}).get('template_hash') == sha256_file(tpl):
            continue
        # A rule for one table cell does not cover another occurrence of the
        # same field. Compare real paragraph locations, not just field names.
        known = {r['field'] for r in rows}
        for aliases in SYNONYMS.values():
            known.update(aliases)
        fields = {r['field']: {} for r in rows}
        for candidate in extract_labels(tpl, known_labels=known):
            ranked = candidates_for(candidate['text'], fields)
            if not ranked or ranked[0]['score'] < AUTO_APPLY_THRESHOLD:
                continue
            if len(ranked) > 1 and ranked[0]['score'] == ranked[1]['score']:
                continue
            field = ranked[0]['field']
            if any(r['field'] == field and r['target'].get('only_first') for r in rows):
                continue
            try:
                missed = positions(candidate['target']) - covered.get(field, set())
                if missed:
                    errors.append({'template': tpl.name, 'field': field,
                                   'location': candidate['location'], 'target': candidate['target'],
                                   'reason': '同一字段还有明确填写位置未被当前规则覆盖，请核对具体位置。'})
            except (OfficeKitError, IndexError, KeyError, ValueError, TypeError) as exc:
                errors.append({'template': tpl.name, 'field': field, 'reason': str(exc)})
    return errors


def fill_plan_signature(plan: dict) -> list[tuple]:
    """Only business-relevant parts: detect rule/value changes before writing."""
    keys = ("template", "template_sha256", "template_rule_id", "field", "label", "target", "value",
            "decision", "ask_reason", "entity_id", "provenance", "is_required")
    return sorted(tuple(json.dumps(row.get(key), ensure_ascii=False, sort_keys=True)
                        for key in keys)
                  for row in plan["rows"] if row.get("kind") == "slot")


def parse_fill_selection(select: str | None, apply_all: bool = False,
                         new: list[str] | None = None,
                         blank: list[str] | None = None,
                         use: list[str] | None = None) -> dict:
    """⚙ 目标侧**三选**：接受建议 / 输入新值 / 留空（`27 §5.1` 乙表）。

    外加一条（T-25）：`--use N=<主体名或#编号>` —— 这一格**用哪个主体的值**。
    只在"库里这个键属于多个主体、字段名又没指明"时才需要。
    """
    actions: dict[int, dict] = {}

    def take(num, action, value=None):
        try:
            n = int(str(num).split("=")[0].strip())
        except ValueError as exc:
            raise OfficeKitError("勾选编号必须是数字：%r" % num) from exc
        if n in actions:
            raise OfficeKitError("编号 %d 被勾了两次（%s 与 %s）"
                                 % (n, actions[n]["action"], action))
        actions[n] = {"action": action, "value": value}

    for token in (select or "").replace("，", ",").split(","):
        if token.strip():
            take(token, "accept")
    for token in new or []:
        if "=" not in str(token):
            raise OfficeKitError("--new 要写成 编号=值，例如 --new 3=广东众森实业发展有限公司")
        num, value = str(token).split("=", 1)
        take(num, "new", value)
    for token in blank or []:
        take(token, "blank")
    for token in use or []:
        if "=" not in str(token):
            raise OfficeKitError("--use 要写成 编号=主体名，例如 --use 3=广东众森实业发展有限公司")
        num, who = str(token).split("=", 1)
        take(num, "use", who.strip())
    return {"actions": actions, "apply_all": bool(apply_all),
            "chosen": sorted(actions)}


def _guard_fill(plan: dict, selection: dict) -> tuple[list[dict], list[dict]]:
    slots = {r["n"]: r for r in plan["rows"] if r["kind"] == "slot"}
    problems: list[dict] = []
    todo: dict[int, dict] = {}

    for n in selection["chosen"]:
        if n not in slots:
            problems.append({"n": n, "why": "计划里没有这个编号"})
    if selection["apply_all"]:
        for r in plan["rows"]:
            if r["kind"] != "slot":
                continue
            if r["decision"] == "auto":
                todo[r["n"]] = {"action": "accept", "value": None}
            elif r["decision"] == "ask" and r["n"] not in selection["actions"]:
                problems.append({"n": r["n"], "why": "全选不含它：" + r["ask_reason"]
                                 + "，请单独写编号勾"})
            # empty 不勾（没值可填）
    for n, act in selection["actions"].items():
        if n in slots:
            todo[n] = act

    accepted: list[dict] = []
    for n in sorted(todo):
        r = slots[n]
        act = todo[n]
        if act["action"] in ("accept", "new") and r["decision"] == "empty" \
                and act["action"] != "new":
            problems.append({"n": n, "why": "源文件里没有这个值，选不了「接受」——"
                                            "要么给它一个新值（--new %d=值），要么留空" % n})
            continue
        # 🔴 歧义格：库里这个键属于多个主体 → **必须指明用谁的**
        if act["action"] == "use" and not r.get("ambiguous"):
            # 没有歧义的格子不用 --use —— 说明白，别让它默默生效
            problems.append({"n": n, "why":
                             "这个格子库里只有一条值（%s），没有歧义，不用 --use；"
                             "直接 `--select %d` 就行" % (r.get("entity_name"), n)})
            continue
        if r.get("ambiguous") and act["action"] in ("accept", "use"):
            if act["action"] != "use":
                problems.append({"n": n, "why":
                                 "库里这个键有 %d 个主体的值（%s），字段名没指明是谁的——"
                                 "请用 `--use %d=<主体名或#编号>`，或 `--new %d=值` 自己给一个"
                                 % (len(r.get("candidates") or []),
                                    "、".join(c["entity_name"]
                                              for c in (r.get("candidates") or [])[:3]),
                                    n, n)})
                continue
            hit = _match_candidate(r.get("candidates") or [], act.get("value") or "")
            if hit is None:
                problems.append({"n": n, "why":
                                 "指明的「%s」不在这个键的候选里（候选：%s）"
                                 % (act.get("value"),
                                    "、".join("#%s %s" % (c["entity_id"], c["entity_name"])
                                              for c in (r.get("candidates") or [])))})
                continue
            accepted.append({**r, "action": "accept",
                             "value": hit["value"], "entity_id": hit["entity_id"],
                             "entity_name": hit["entity_name"], "_picked": True})
            continue
        if act["action"] == "new":
            # A custom target value is a local override, not evidence that it
            # belongs to one of the source candidates.  Cross-subject checking
            # may use an already unique template subject, otherwise it must say
            # that the value has no source entity.
            new_eid = r.get("subject_eid")
            accepted.append({**r, **act, "entity_id": new_eid,
                             "entity_name": r.get("subject_label") if new_eid is not None else "（未归属）",
                             "_target_override": True})
            continue
        accepted.append({**r, **act})
    return accepted, problems


def _match_candidate(cands: list[dict], who: str) -> dict | None:
    """把 `--use` 给的主体名 / `#编号` 对上候选。**对不上就返回 None，不猜。**"""
    w = str(who).strip().lstrip("#")
    for c in cands:
        if w and (str(c["entity_id"]) == w or c["entity_name"] == who
                  or c["entity_name"].strip() == w):
            return c
    return None


def _role_claimed_by_field(field: str) -> str | None:
    """这个字段名**声称**值的归属主体扮演什么角色。

    用的是 `schema_v2.map_role` 那一套**关键词自上而下、先命中先算**的判法，
    与迁移同一口径，免得两处各写一套：

    * `借款人法定代表人` → **法定代表人**（不是"借款人"——值是个人，属人身份）；
    * `保证人名称` → 保证人；`借款人名称` → 借款人；
    * `流贷合同号` → 不声称任何角色（返回 None，不参与本项核对）。
    """
    from .schema_v2 import map_role

    return map_role(field)


def _check_cross_subject(store: StoreV2, rows: list[dict]) -> dict:
    """红线一的**本期口径**：角色一致性 + 无归属统计。

    * 字段名声称了某个角色（如 `保证人名称`、`借款人法定代表人`）→
      这条值的归属主体**必须真的扮演那个角色**；
    * 值没有归属 → 记进"无法核对"，**不假装核对过**。
    """
    violations: list[dict] = []
    unchecked: list[dict] = []
    checked = 0
    # **这份产物允许出现哪些主体**（T-25 兜底）：
    #   · 本产物主体（借款人那张表就是借款人）；
    #   · 加上**本表任何格子声称的角色**所对应的主体
    #     —— 借款人表上写"保证人是谁"是合法的，不能一刀切。
    # 一个值的主体**不在这个集合里** → 那就是串了别人的资料。
    allowed: set[int] = set()
    if rows:
        subj = rows[0].get("subject_eid")
        if subj is not None:
            allowed.add(int(subj))
        for r in rows:
            c = _role_claimed_by_field(r["field"])
            if c:
                e = store.role_entity(c)
                if e is not None:
                    allowed.add(int(e))
    for r in rows:
        field = r["field"]
        eid = r.get("entity_id")
        claimed = _role_claimed_by_field(field)
        if r.get('qualified_source') and r.get('relation_owner_eid') is not None:
            # This association was read from the source's explicit company ->
            # representative relationship, not inferred from a model's words.
            owner = int(r['relation_owner_eid'])
            company_role = next((role for role in ('借款人', '保证人') if field.startswith(role)), None)
            if company_role and company_role in store.roles_of(owner):
                checked += 1
                continue
        if claimed is None:
            # 不声称角色的字段（`联系电话`、`开户行`…）**不再直接跳过**：
            # 只看一件事——它的值属不属于这份产物。
            #
            # ⚠️ **只在"库里这个键本来就有多个候选"时才判**（T-25 收窄过一次）：
            #    如果这个键**只有一条**，那它的归属就是"入库时怎么归的"，
            #    跟"填表时挑错了"是两件事——拿它拦填报会**误杀**。
            #    实测例子：「法定代表人身份证明书」这张表上，
            #    借款人公司的 `联系电话` 会被误判成"不属于这份产物"
            #    （因为表上唯一的角色字段是 `借款人法定代表人`，主体被认成了顾红军）。
            #    那种情况该由 `db-report` 的"待归属"去说，不该拦死填报。
            if eid is None:
                unchecked.append({"field": field, "template": r["template"],
                                  "why": "这条值没有归属主体，无法核对"})
                continue
            multi = len(store.current_rows(field)) > 1
            if multi and allowed and int(eid) not in allowed:
                violations.append({
                    "field": field, "template": r["template"], "entity_id": eid,
                    "entity_name": store.entity_label(int(eid)),
                    "why": "「%s」这一格的值来自「%s」，**不属于这份产物**"
                           "（本产物相关主体：%s）——跨主体串数据"
                           % (field, store.entity_label(int(eid)),
                              "、".join(sorted(store.entity_label(x) for x in allowed)))})
            else:
                checked += 1
            continue
        if eid is None:
            unchecked.append({"field": field, "template": r["template"],
                              "why": "这条值没有归属主体，无法核对"})
            continue
        roles = store.roles_of(int(eid))
        if claimed in roles:
            checked += 1
        elif not roles:
            # 这个主体**一个角色都没记**（源文件没明写、你还没确认）。
            # 那是 🟡「角色未确认」，**不是**「跨主体串数据」——两件事不能混，
            # 否则每份"源文件没写角色"的材料都会被误判成红线一违规。
            unchecked.append({
                "field": field, "template": r["template"], "entity_id": eid,
                "entity_name": store.entity_label(int(eid)),
                "why": "「%s」这一格的值来自「%s」，但它在库里**还没有任何角色记录**"
                       "（源文件没明写、也还没人工确认）——列进「无法核对」，"
                       "**不算违规**" % (field, store.entity_label(int(eid)))})
        else:
            violations.append({
                "field": field, "template": r["template"], "entity_id": eid,
                "entity_name": store.entity_label(int(eid)),
                "why": "「%s」这一格的值来自「%s」，但它在库里没有「%s」这个身份（只有 %s）"
                       % (field, store.entity_label(int(eid)), claimed,
                          "、".join(roles) or "无")})
    return {"ok": not violations, "checked": checked, "violations": violations,
            "unchecked": unchecked,
            "note": "本期口径 = 字段名声称的角色与值归属主体的实际角色必须一致；"
                    "完整红线一（逐值反查产物主体）见 28 §6"}


def _check_no_fabrication(rows: list[dict]) -> dict:
    """红线二：每个值都必须说得出它是哪来的（source / computed / user）。"""
    bad = []
    for r in rows:
        kind = r.get("_source_kind")
        if kind not in ("source", "computed", "user"):
            bad.append({"field": r["field"], "template": r["template"],
                        "source_kind": kind})
    return {"ok": not bad, "bad": bad}


def cmd_db_fill(args) -> Result:
    """⑩⑪⑫⑬：预演 → 校验 → 保真填充 → 签核。**一次运行可以出多份产物。**"""
    res = Result("db-fill")
    wr, store = _open(args)
    engine_name = (getattr(args, "engine", None) or "xml").lower()
    batch_no = getattr(args, "batch", None) or store.current_batch()
    if not batch_no:
        store.close()
        raise OfficeKitError("库里还没有批次——先跑 db-ingest")
    # 报告落 `<工作区>\out\<批次号>\_报告\db-fill\`；产物另走主体目录
    out = _rdir(args, "db-fill", batch_no)

    # 识别码、证件号码是可选资料，不作为填报前置条件。
    gate = _gate_identity(store)

    templates = resolve_inputs(args.input)
    plan_file = getattr(args, "plan", None)

    # ---------- ⑩ 预演 ----------
    if plan_file:
        p = Path(plan_file)
        if not p.exists():
            store.close()
            raise OfficeKitError("预演清单不存在：%s" % p)
        plan = json.loads(p.read_text(encoding="utf-8"))
        if plan.get("batch_no") != batch_no:
            store.close()
            raise OfficeKitError("业务批次已经变化，请重新预演本批目标文件并确认；不能沿用旧批次的答案。")
        if plan.get("db_path") and Path(plan["db_path"]).resolve() != Path(args._db).resolve():
            store.close()
            raise OfficeKitError("预演清单属于另一个办公数据库；请在当前工作区重新预演。")
        run_id = plan["run_id"]
    else:
        run_id = _next_run_id(store, batch_no)
        from .rule_pack import reuse_verified
        reused = []
        reuse_conflicts = []
        for tpl in templates:
            reused_result = reuse_verified(store, tpl, batch_no)
            if reused_result and reused_result["added"]:
                reused.append(tpl.name)
            if reused_result and reused_result["conflicts_skipped"]:
                reuse_conflicts.append(tpl.name)
        if reused:
            res.warn("已自动复用 %d 份内容完全相同、此前经使用者复核的模板位置：%s"
                     % (len(reused), "、".join(reused[:4])))
        if reuse_conflicts:
            res.warn("以下模板在本工作区有不同的填写规则，自动复用未覆盖它们：%s"
                     % "、".join(reuse_conflicts[:4]))
        plan = build_fill_plan(store, templates, batch_no=batch_no, run_id=run_id)

    plan["db_path"] = str(Path(args._db).resolve())
    plan["template_files"] = [str(t.resolve()) for t in templates]

    # 运行行（kind='run'）只建一次；它承载"执行前快照"。
    if store.conn.execute("SELECT id FROM fill_op WHERE run_id=? AND kind='run'",
                          (run_id,)).fetchone() is None:
        from .rule_pack import rules_digest
        preflight = {"batch_no": batch_no, "run_id": run_id,
                     "tool_fingerprint": _tool_fp(),
                     "templates": [
                         {"name": t.name, "sha256": sha256_file(t),
                          "rules_sha256": rules_digest(store, t),
                          "structure_ok": _template_structure(t)[0],
                          "structure_note": _template_structure(t)[2],
                          "slots": sum(1 for r in plan["rows"]
                                       if r["kind"] == "slot" and r["template"] == t.name)}
                         for t in templates]}
        store.start_run(batch_no, run_id, preflight=preflight)
        store.set_run_status(run_id, "running", "dry_run_only")

    selecting = bool(getattr(args, "confirmed", False) or getattr(args, "select", None) or getattr(args, "apply_all", False)
                     or getattr(args, "new", None) or getattr(args, "blank", None)
                     or getattr(args, "use", None))

    # In DSH the fill-card database is the authority, even if the model uses
    # legacy selection flags. A plain --plan also follows this default outside
    # DSH; --confirmed remains a backward-compatible spelling, not a requirement.
    database_mode = bool(getattr(args, "confirmed", False) or
                         (plan_file and (os.environ.get("DSH_SESSION_ID") or not selecting)))
    if os.environ.get("DSH_SESSION_ID") and not plan_file:
        selecting = False
    saved_selection = None
    if database_mode:
        from .fill_decisions import saved_choices
        choices, missing = saved_choices(plan)
        if missing:
            res.data.update({"status": "awaiting_confirmation", "blocked": True,
                             "plan_file": str(Path(plan_file).resolve()) if plan_file else None,
                             "missing_confirmation_count": len(missing),
                             "next_tool": "office_fill_review",
                             "next_tool_arguments": {"plan_file": str(Path(plan_file).resolve())}
                             if plan_file else {}})
            res.warn("需要确认的内容尚未齐全。请由系统调用 office_fill_review 弹窗；"
                     "回答会自动保存，随后继续本次填报，不需要用户提供操作口令。")
            store.close()
            return res
        saved_selection = parse_fill_selection(choices["select"], True,
                                                choices["new"], choices["blank"], choices["use"])
        selecting = True

    if not selecting:
        # Always inspect the actual templates and current rules. An old or
        # absent db-propose report cannot silently disable this check.
        if not plan_file:
            plan["uncovered_xlsx_labels"] = current_xlsx_gaps(store, templates, plan)
            plan["docx_target_errors"] = current_docx_target_errors(templates, plan)
            plan["templates_without_rules"] = [r["template"] for r in plan["rows"]
                                               if r["kind"] == "template_note"]
        store.event("dryrun_snapshot", batch_no=batch_no, run_id=run_id,
                    payload={"templates": [t.name for t in templates],
                             "counts": plan["counts"]})
        md = render_fill_plan(store, plan)
        p_md = write_text(unique_path(out / "fill_plan.md"), md)
        p_json = write_text(unique_path(out / "fill_plan.json"),
                            json.dumps(plan, ensure_ascii=False, indent=2))
        res.add_artifact(p_md, "预演清单（给人看：哪个格子填什么、为什么问你）")
        res.add_artifact(p_json, "同一份预演清单（执行时用 --plan 传回来）")
        _warn_outside(res, args, p_md, p_json)
        res.data.update({"blocked": bool(plan.get("uncovered_xlsx_labels") or
                                         plan.get("docx_target_errors") or
                                         plan.get("templates_without_rules")),
                         "batch_no": batch_no, "run_id": run_id,
                         "plan_file": str(p_json), "tool_fingerprint": _tool_fp(),
                         "next_tool": "office_fill_review",
                         "next_tool_arguments": {"plan_file": str(p_json)},
                         "uncovered_xlsx_labels": plan.get("uncovered_xlsx_labels", []),
                         "docx_target_errors": plan.get("docx_target_errors", []),
                         "templates_without_rules": plan.get("templates_without_rules", []),
                         **plan["counts"], "summary": store.summary()})
        c = plan["counts"]
        res.warn("预演完成：%d 个格子，其中自动填 %d、要你点头 %d、源里没有 %d。"
                 "还没有产出文件。系统下一步调用 office_fill_review，自动查库并弹出尚未确认的项；"
                 "随后带 --plan 执行全部目标文件。用户无需提供特殊提示词。"
                 % (c["slots"], c["auto"], c["ask"], c["empty"]))
        gaps = plan.get("uncovered_xlsx_labels", [])
        if gaps:
            res.warn("⛔ %d 个 Excel 标签匹配到现有字段，却没有对应填写规则：%s。"
                     "此预演不能进入确认或填报；先核对并补规则。"
                     % (len(gaps), "；".join("%s / %s" % (g["template"], g["label"])
                                            for g in gaps[:4])))
        target_errors = plan.get("docx_target_errors", [])
        if target_errors:
            res.warn("⛔ %d 处 Word 填写位置无效：%s。先修规则再请使用者确认。" %
                     (len(target_errors), "；".join("%s / %s：%s" %
                      (e["template"], e["field"], e["reason"]) for e in target_errors[:4])))
        no_rules = plan.get("templates_without_rules", [])
        if no_rules:
            res.warn("⛔ %d 份目标模板没有任何填写规则：%s。先补规则再请使用者确认。" %
                     (len(no_rules), "、".join(no_rules[:4])))
        store.close()
        return res

    # ---------- ⑪ 校验 + ⑫ 填充 ----------
    if saved_selection is not None:
        selection = saved_selection
    else:
        selection = parse_fill_selection(getattr(args, "select", None),
                                         getattr(args, "apply_all", False),
                                         getattr(args, "new", None),
                                         getattr(args, "blank", None),
                                         getattr(args, "use", None))
    current_plan = build_fill_plan(store, templates, batch_no=batch_no, run_id=run_id)
    if fill_plan_signature(plan) != fill_plan_signature(current_plan):
        store.close()
        raise OfficeKitError("模板位置、填写规则或源数据在预演后变了，已停止。请重新预演；未变化的回答会沿用。")
    no_rules = [r["template"] for r in current_plan["rows"] if r["kind"] == "template_note"]
    if no_rules:
        store.close()
        raise OfficeKitError("%d 份目标模板没有填写规则，已停止填报：%s。请先补规则并重新预演。" %
                             (len(no_rules), "、".join(no_rules[:4])))
    gaps = current_xlsx_gaps(store, templates, plan)
    if gaps:
        store.close()
        raise OfficeKitError("Excel 模板有 %d 处明确标签缺少填写规则，已停止填报：%s。"
                             "请先补规则、重新预演，再请使用者确认。" %
                             (len(gaps), "；".join("%s / %s" % (g["template"], g["label"])
                                                 for g in gaps[:4])))
    target_errors = current_docx_target_errors(templates, plan)
    if target_errors:
        store.close()
        raise OfficeKitError("Word 模板有 %d 处填写位置无效，已停止填报：%s。请先修规则并重新预演。" %
                             (len(target_errors), "；".join("%s / %s：%s" %
                              (e["template"], e["field"], e["reason"]) for e in target_errors[:4])))
    accepted, problems = _guard_fill(plan, selection)
    chosen = {a["n"]: a for a in accepted}
    slots = [r for r in plan["rows"] if r["kind"] == "slot"]

    # 目标侧三选：接受 / 输入新值 / 留空
    writes: list[dict] = []
    blanked: list[dict] = []
    for r in slots:
        act = chosen.get(r["n"])
        if act is None:
            if r["decision"] == "empty":
                # 源文件里确实没有 → 这一格留空，但**要记账、要进待确认清单**，
                # 不能因为"没值"就悄悄跳过（否则缺失项永远没人看见）。
                store.add_review(batch_no, "field", "value_missing_in_source", run_id=run_id,
                                 entity_id=r["entity_id"], template_id=r["template_id"],
                                 field=r["field"], label=r["label"],
                                 candidates={"required": r["is_required"],
                                             "provenance": r["provenance"]})
                store.record_field(run_id, batch_no, r["field"], "missing",
                                   entity_id=r["entity_id"], template_id=r["template_id"],
                                   template_rule_id=r["template_rule_id"],
                                   detail="源文件里没有这个值（缺失留空）")
                continue
            # 没被勾 → 这一格不填，进待确认清单（**不是**悄悄跳过）
            store.add_review(batch_no, "field", "not_selected", run_id=run_id,
                             entity_id=r["entity_id"], template_id=r["template_id"],
                             field=r["field"], label=r["label"],
                             candidates={"value": r["value"], "decision": r["decision"],
                                         "why": r["ask_reason"]})
            store.record_field(run_id, batch_no, r["field"], "skipped",
                               entity_id=r["entity_id"], template_id=r["template_id"],
                               template_rule_id=r["template_rule_id"],
                               detail="使用者没有勾选这一格")
            continue
        if act["action"] == "blank":
            blanked.append(r)
            store.add_review(batch_no, "field", "blanked_by_user", run_id=run_id,
                             entity_id=r["entity_id"], template_id=r["template_id"],
                             field=r["field"], label=r["label"], candidates=r["value"])
            store.record_field(run_id, batch_no, r["field"], "skipped",
                               entity_id=r["entity_id"], template_id=r["template_id"],
                               template_rule_id=r["template_rule_id"],
                               detail="使用者选择留空")
            continue
        if act["action"] == "new":
            # A target-cell correction belongs to this output only. Updating
            # the shared fact here would silently change every other form and
            # force new previews/questions after the user already answered.
            writes.append({**r, "value": act["value"], "entity_id": act.get("entity_id"),
                           "entity_name": act.get("entity_name"), "_source_kind": "user",
                           "_fact_id": None, "_target_override": True})
        elif act.get("_picked"):
            # `--use` 指定的那个主体的值：值本身来自库，来源类型也照库里的
            writes.append({**r, "value": act["value"], "entity_id": act["entity_id"],
                           "entity_name": act.get("entity_name"),
                           "_source_kind": r.get("source_kind") or "source",
                           "_fact_id": act.get("_fact_id")})
        else:
            writes.append({**r, "_source_kind": r.get("source_kind") or "source",
                           "_fact_id": None})

    # 红线二：不许无中生有
    fabric = _check_no_fabrication(writes)
    if not fabric["ok"]:
        store.event("fill_blocked", batch_no=batch_no, run_id=run_id,
                    payload={"why": "value_without_origin", "bad": fabric["bad"]})
        store.set_run_status(run_id, "blocked", "value_without_origin")
        store.close()
        raise OfficeKitError("🔴 红线二：有值说不出来源，拒绝填充：%s" % fabric["bad"])

    # 红线一（本期口径）：角色前缀一致性
    cross = _check_cross_subject(store, writes)
    if not cross["ok"]:
        store.event("fill_blocked", batch_no=batch_no, run_id=run_id,
                    payload={"why": "cross_subject", **cross})
        store.set_run_status(run_id, "blocked", "cross_subject")
        text = render_cross_subject(cross, batch_no)
        p = write_text(out / "跨主体被拦下.md", text)
        res.add_artifact(p, "红线一：跨主体串数据被拦下")
        res.data.update({"blocked": True, "reason": "cross_subject", "cross": cross,
                         "batch_no": batch_no, "run_id": run_id})
        res.warn("🔴 红线一：%d 个格子的值来自不该出现在这里的角色，**已停止填充**"
                 % len(cross["violations"]))
        store.close()
        return res

    by_template: dict[str, list[dict]] = {}
    for r in writes:
        by_template.setdefault(r["template"], []).append(r)

    results: list[dict] = []
    for tpl in templates:
        rows = by_template.get(tpl.name) or []
        if not rows:
            res.warn("%s：本次没有要填的格子，未产出文件" % tpl.name)
            continue
        struct_ok, kind, struct_note = _template_structure(tpl)
        tid = store.register_template(tpl, batch_no, doc_kind=kind, structure_ok=struct_ok,
                                      structure_note=struct_note)
        # 产物按**主体**分目录：`<工作区>\out\<批次号>\<主体名>\`（26 §2.9）。
        # "这份表单是谁的材料"取值顺序（**只用来决定放哪个文件夹**，不是业务断言）：
        #   借款人 > 保证人 > 法定代表人 > 唯一主体 > 多主体 > 未归属
        # 借款人那张表上通常还要写保证人是谁，所以不能按"值里有几个主体"来分，
        # 否则每份表单都会变成「多主体」。
        subject = _artifact_subject(store, rows)
        art_dir = WR.artifact_dir_for(args, batch_no, subject)
        _warn_outside(res, args, art_dir)
        n_ok, n_fail, dst, deliverable = _fill_one_template(
            store, tpl, rows, run_id, batch_no, tid, engine_name, art_dir, res)
        results.append({"template": tpl.name, "filled": n_ok, "failed": n_fail,
                        "slots": len(rows), "subject": subject,
                        "deliverable": deliverable,
                        "output": str(dst), "outputs": [str(dst)]})

    total_ok = sum(r["filled"] for r in results)
    has_rejected = any(not r["deliverable"] for r in results)
    # ---- 签核（T-26）----------------------------------------------------
    # 🔴 **两种"签名"必须分开**：
    #   · `--sign`           = **代理按你说的填的名字**。它是**声明**，不是证据。
    #   · `--sign-confirmed` = **使用者在对话里确认过**，而且要**带上他的原话**。
    #                         原话可以拿去 DSH 的会话记录里对——**名字不值钱，能对上原话才值钱**。
    # 交付门禁四① **只认后者**。
    sign_kind = None
    if has_rejected and (getattr(args, "sign_confirmed", None)
                         or getattr(args, "sign", None)):
        res.warn("本轮有未通过填充或保真检查的草稿，未记录整批签核；先修复失败格。")
    elif getattr(args, "sign_confirmed", None):
        quote = (getattr(args, "quote", None) or "").strip()
        if not quote:
            store.close()
            raise OfficeKitError(
                "--sign-confirmed 必须同时给 --quote \"使用者的原话\"——\n"
                "  没有原话，这个名字和 --sign 一样，只是程序写进去的一串字。")
        actor = args.sign_confirmed
        sign_kind = "confirmed-by-user"
        n = store.signoff(run_id, decision="approve", actor=actor, scope="batch",
                          fingerprint="%s|%s" % (sign_kind, _run_fingerprint(store, run_id)))
        store.event("signoff_confirmed", batch_no=batch_no, run_id=run_id, target="签核",
                    payload={"actor": actor, "kind": sign_kind, "quote": quote,
                             "session_ref": os.environ.get("DSH_SESSION_ID"),
                             "note": "使用者在对话里确认过；原话可拿去 DSH 会话记录对照"},
                    actor=actor)
        res.data["signed_off"] = {"actor": actor, "kind": sign_kind, "quote": quote,
                                  "artifacts": n}
    elif getattr(args, "sign", None):
        sign_kind = "declared-by-agent"
        n = store.signoff(run_id, decision="approve", actor=args.sign, scope="batch",
                          fingerprint="%s|%s" % (sign_kind, _run_fingerprint(store, run_id)))
        store.event("signoff_declared", batch_no=batch_no, run_id=run_id, target="签核",
                    payload={"actor": args.sign, "kind": sign_kind,
                             "note": "⚠️ 这是**代理代填**的名字，**不等于使用者检查过成品**；"
                                     "要真签核请用 --sign-confirmed 并附使用者原话"},
                    actor=args.sign)
        res.data["signed_off"] = {"actor": args.sign, "kind": sign_kind, "artifacts": n}
    store.set_run_status(run_id, "rejected" if has_rejected else
                         "ok" if results else "missing",
                         "有填充失败或保真检查失败" if has_rejected else
                         "" if results else "没有产物")
    store.event("fill_executed", batch_no=batch_no, run_id=run_id,
                payload={"results": results, "cross_subject": cross["checked"],
                         "unchecked": len(cross["unchecked"]),
                         "problems": problems, "blanked": len(blanked),
                         # **这个版本的工具做的**——分诊与复现都靠它
                         "tool_fingerprint": _tool_fp()})
    store.event("verify_performed", batch_no=batch_no, run_id=run_id,
                payload={"cross_subject": cross, "no_fabrication": fabric})
    if sign_kind == "confirmed-by-user" and not has_rejected:
        from .rule_pack import publish_confirmed
        for tpl in templates:
            if any(r["template"] == tpl.name and r["deliverable"] for r in results):
                try:
                    publish_confirmed(store, tpl)
                except (OfficeKitError, OSError, ValueError) as exc:
                    res.warn("模板自动复用目录未更新：%s" % exc)

    res.data.update({
        "blocked": has_rejected, "batch_no": batch_no, "run_id": run_id,
        "templates": results, "total_filled": total_ok,
        "problems": problems,
        "blanked_by_user": [r["field"] for r in blanked],
        # 因为"这个主体的角色没确认"而没填的格子（🟡 提示级，不是硬拦）
        "role_unconfirmed_fields": sorted({
            r["field"] for r in slots
            if r["decision"] == "ask" and "角色" in (r["ask_reason"] or "")}),
        "cross_subject": cross,
        # ⚠️ **只有"使用者确认"才算签核**。`--sign`（代理代填）**仍然是未签核**——
        # 这一条写错过一次：原来写成 `sign_kind is None`，于是"代理代填"被当成了已签核，
        # 门禁四① 又被打开了。是 selftest_pipeline_v2 的第 [13] 段抓出来的。
        "unsigned": sign_kind != "confirmed-by-user",
        "signoff_kind": sign_kind,
        "tool_fingerprint": _tool_fp(),
        "reviews_open": len(store.open_reviews()),
        "summary": store.summary(),
    })
    md = render_fill_result(store, plan, results, cross, run_id, batch_no,
                            signed=(sign_kind == "confirmed-by-user"),
                            sign_kind=sign_kind,
                            sign_actor=(res.data.get("signed_off") or {}).get("actor"))
    p_md = write_text(out / "填报结果.md", md)
    _warn_outside(res, args, p_md)
    res.add_artifact(p_md, "填报结果（逐格结果 + 保真证明 + 门禁判定）")
    if problems:
        res.warn("%d 个编号被拒绝（%s）" % (len(problems), problems[0]["why"][:60]))
    if sign_kind is None:
        res.warn("🔴 本次产物**未签核**（差交付门禁四①）——未签核不许交付。"
                 "使用者本人复核后，用 `--sign-confirmed <姓名> --quote \"他的原话\"` 记录签核")
    elif sign_kind == "declared-by-agent":
        res.warn("⚠️ 签核记的是**「代理代填」**（`--sign %s`）——"
                 "**这不等于使用者检查过成品**，交付门禁四① **不认它**。"
                 "要真签核请用 `--sign-confirmed <姓名> --quote \"使用者的原话\"`"
                 % res.data["signed_off"]["actor"])
    store.close()
    return res


def _next_run_id(store: StoreV2, batch_no: str) -> str:
    """一次运行一个号：`<批次号>-R<序号>`（`27 §3.3` 的 `run_id` 拼法）。"""
    n = store.conn.execute("SELECT COUNT(*) FROM fill_op WHERE batch_no=? AND kind='run'",
                           (batch_no,)).fetchone()[0]
    while True:
        n += 1
        cand = "%s-R%d" % (batch_no, n)
        if store.conn.execute("SELECT id FROM fill_op WHERE run_id=?", (cand,)).fetchone() is None:
            return cand


def _run_fingerprint(store: StoreV2, run_id: str) -> str:
    import hashlib

    h = hashlib.sha256()
    for r in store.conn.execute("SELECT artifact_path, artifact_sha256 FROM fill_op"
                                " WHERE run_id=? AND kind='artifact' ORDER BY id", (run_id,)):
        h.update(("%s|%s\n" % (r[0], r[1])).encode("utf-8"))
    return h.hexdigest()


def _artifact_subject(store: StoreV2, rows: list[dict]) -> str:
    """这份表单是**谁**的材料（用来决定产物放哪个文件夹）。

    顺序：**借款人 > 保证人 > 法定代表人 > 唯一主体 > 多主体 > 未归属**。
    ⚠️ 这**只是目录名**，不是业务断言——所以顺序是固定的、写进报告里，
    不随"这次恰好填了哪几个值"漂移（否则同一份表单两次运行会落进不同目录）。
    """
    eids = {r["entity_id"] for r in rows if r.get("entity_id") is not None}
    for role in ("借款人", "保证人", "法定代表人"):
        for eid in sorted(eids):
            if role in store.roles_of(int(eid)):
                return store.entity_label(int(eid))
    if len(eids) == 1:
        return store.entity_label(next(iter(eids)))
    return "多主体" if eids else "未归属"


def _fill_one_template(store: StoreV2, tpl: Path, rows: list[dict], run_id: str,
                       batch_no: str, tid: int, engine_name: str, out: Path,
                       res: Result) -> tuple[int, int, Path, bool]:
    """填一份并检查原格式；失败的草稿不列为可交付产物。"""
    import docx

    book = engine = document = None
    if tpl.suffix.lower() in (".xlsx", ".xlsm"):
        book = doc_fill.XlsxEngine(tpl)
    else:
        if engine_name == "xml":
            try:
                engine = doc_fill.XmlEngine(tpl)
            except OfficeKitError as exc:
                res.warn("%s：XML 引擎不可用（%s），改用 python-docx" % (tpl.name, exc))
        document = engine.document if engine else docx.Document(str(tpl))

    n_ok = n_fail = 0
    for r in rows:
        field, value, target = r["field"], r["value"], r["target"]
        try:
            if target.get("kind") == "xlsx_cell":
                if book is None:
                    raise OfficeKitError("目标不是工作簿")
                book.fill_xlsx_cell(target, value)
                ok = True
            elif target.get("kind") == "cell":
                if engine:
                    ok = engine.fill_cell(target, value)
                else:
                    doc_fill._apply_cell(document, target, _flat(store), value)
                    ok = True
            else:
                if engine:
                    hits = engine.find_anchor_hits(target)
                    if target.get("only_first") and len(hits) > 1:
                        hits = hits[:1]
                    if not hits:
                        raise OfficeKitError("label_not_found")
                    if len(hits) > int(target.get("max_matches", 1)):
                        raise OfficeKitError("ambiguous:%d" % len(hits))
                    for par, s, e in hits:
                        if not engine.fill_span(par, s, e, value):
                            raise OfficeKitError("write_failed")
                    ok = True
                else:
                    locs, status = doc_fill._apply_anchor(document, target, _flat(store), value)
                    ok = status == "filled"
                    if status == "ambiguous":
                        raise OfficeKitError("ambiguous:%d" % len(locs))
                    if status == "label_not_found":
                        raise OfficeKitError("label_not_found")
            if not ok:
                raise OfficeKitError("write_failed")
            n_ok += 1
            store.record_field(run_id, batch_no, field, "ok", entity_id=r["entity_id"],
                               template_id=tid, template_rule_id=r["template_rule_id"],
                               fact_id=r.get("_fact_id"), value=str(value),
                               detail="使用者仅为本目标位置改值" if r.get("_target_override") else "")
        except OfficeKitError as exc:
            detail = str(exc)
            reason = ("ambiguous_multiple_matches" if detail.startswith("ambiguous")
                      else "label_not_found" if detail == "label_not_found"
                      else "rule_failed")
            n_fail += 1
            store.add_review(batch_no, "field", reason, run_id=run_id,
                             entity_id=r["entity_id"], template_id=tid,
                             field=field, label=r["label"], candidates=detail)
            store.record_field(run_id, batch_no, field, "skipped", entity_id=r["entity_id"],
                               template_id=tid, template_rule_id=r["template_rule_id"],
                               fact_id=r.get("_fact_id"), value=str(value), detail=detail)

    suffix = ".xlsx" if book is not None else ".docx"
    dst = unique_path(out / ("%s_已填写%s" % (safe_stem(tpl.stem), suffix)))
    dst.parent.mkdir(parents=True, exist_ok=True)
    if book is not None:
        book.save(dst)
    elif engine:
        # **必须走引擎保存**：这里调 document.save() 会用 python-docx 的 DOM 重新序列化，
        # 把我们做的字节级拼接**全部丢掉**——那时运行会报成功、文件却还是空的。
        engine.save(dst)
    else:
        doc_fill.save_pydocx_with_repair(document, tpl, dst, res)

    # Word 看包内结构；Excel 看样式和可见布局。openpyxl 保存会重写 XML，
    # 因此把 Word 的逐部件字节比较套在 Excel 上会误报每一份文件都坏了。
    if book is not None:
        from .xlsx_fidelity import prove_xlsx_fidelity
        fidelity = prove_xlsx_fidelity(tpl, dst, book.edits)
    else:
        fidelity = xml_fill.prove_fidelity(tpl, dst)
    opened, open_note = xml_fill.opens_ok(dst)
    deliverable = bool(n_ok and not n_fail and fidelity["ok"] and opened)
    if not deliverable:
        failure_reason = ("write_failed" if n_fail else
                          "format_changed" if not fidelity["ok"] else
                          "open_failed" if not opened else "no_fields_written")
        store.event("fill_blocked", batch_no=batch_no, run_id=run_id,
                    payload={"why": failure_reason, "template": tpl.name,
                             "fidelity": fidelity, "opened": opened, "note": open_note})
        store.record_artifact(run_id, batch_no, artifact_path=str(dst),
                              artifact_sha256=sha256_file(dst), template_id=tid,
                              opened_ok=opened, status="rejected",
                              detail=("填充失败 %d 格；%s；打开检查=%s"
                                      % (n_fail, fidelity["note"], open_note)))
        res.warn("%s：填充失败 %d 格；%s；打开检查=%s。草稿已留在工作区，"
                 "不作为完成文件交付。" % (tpl.name, n_fail, fidelity["note"], open_note))
    else:
        store.record_artifact(run_id, batch_no, artifact_path=str(dst),
                              artifact_sha256=sha256_file(dst), template_id=tid,
                              opened_ok=True, status="ok", detail=fidelity["note"])
        res.add_artifact(dst, "%s 填充结果（%d 格；%s）"
                         % (tpl.name, n_ok, fidelity["note"]))
    return n_ok, n_fail, dst, deliverable


def _flat(store: StoreV2) -> dict[str, dict]:
    return store.facts_by_key()


# ==========================================================================
# 报告
# ==========================================================================
def render_cross_subject(cross: dict, batch_no: str) -> str:
    L = ["# 🔴 停止填充：红线一（跨主体串数据）", "",
         "| 项 | 值 |", "| --- | --- |", "| 批次 | %s |" % batch_no,
         "| 核对过的格子 | %d |" % cross["checked"],
         "| **违规格子** | **%d** |" % len(cross["violations"]),
         "| 无法核对（值没有归属主体） | %d |" % len(cross["unchecked"]), ""]
    if cross["violations"]:
        L += ["## 违规明细", "", "| 模板 | 字段 | 值的归属主体 | 为什么不行 |",
              "| --- | --- | --- | --- |"]
        for v in cross["violations"]:
            L.append("| %s | `%s` | %s（#%s） | %s |"
                     % (v["template"], v["field"], v["entity_name"], v["entity_id"], v["why"]))
    L += ["", "> 📌 **本期口径**：%s" % cross["note"], ""]
    return "\n".join(L)


def render_fill_plan(store: StoreV2, plan: dict) -> str:
    c = plan["counts"]
    L = ["# 预演清单 / dry-run（**强制**，`27 §6 ⑩`）", "",
         "看了这份清单再决定填什么。**这一趟没有产出任何文件。**", "",
         "| 项 | 数 |", "| --- | --- |",
         "| 批次 | %s |" % plan["batch_no"], "| 本次运行号 | %s |" % plan["run_id"],
         "| 目标模板 | %d 份 |" % c["templates"],
         "| 待填格子 | **%d** |" % c["slots"],
         "| 可自动填（置信度够、非高风险、角色已确认） | %d |" % c["auto"],
         "| **要你逐个点头** | **%d** |" % c["ask"],
         "| 源文件里确实没有 | %d |" % c["empty"], ""]
    gaps = plan.get("uncovered_xlsx_labels", [])
    if gaps:
        L += ["## 先核对这些疑似漏填的 Excel 标签", "",
              "以下标签曾匹配到现有字段，但本次预演没有对应规则。先核对模板与源材料，补规则或说明为何留空，再弹填写确认卡片。", "",
              "| 模板 | 标签 | 建议字段 | 位置 |", "| --- | --- | --- | --- |"]
        for gap in gaps:
            L.append("| %s | %s | %s | %s |" % (
                gap["template"], gap["label"], gap["suggested_field"], gap["location"]))
        L.append("")
    L += ["## 逐格清单", "",
          "| 编号 | 模板 | 标签 | 字段 | 值 | 置信度 | 主体 | 判定 | 为什么 |",
          "| --- | --- | --- | --- | --- | --- | --- | --- | --- |"]
    label = {"auto": "自动填", "ask": "🔴 要你点头", "empty": "⬜ 源里没有"}
    for r in plan["rows"]:
        if r["kind"] != "slot":
            L.append("| — | %s | （无规则：%s） | | | | | | |"
                     % (r["template"], r["note"]))
            continue
        L.append("| %d | %s | %s | `%s` | %s | %.2f | %s | %s | %s |" % (
            r["n"], r["template"], r["label"], r["field"], _cell(r["value"]),
            r["confidence"], r["entity_name"], label[r["decision"]], r["ask_reason"] or "—"))
    L += ["", "## 怎么执行（**目标侧三选**，`27 §5.1` 乙表）", "",
          "```",
          "--select 1,3          # 接受建议",
          "--new 2=你输入的值     # 只改本目标位置，记入填报记录，不改整批共用数据",
          "--blank 4             # 留空（该格空着并留在待确认清单里）",
          "--apply-all           # 全部自动填的那批（**不含**要你点头的格子）",
          "--sign 你的姓名        # 使用者本人签核（干预点 3；不签 = 未签核不许交付）",
          "```", "",
          "> ⚠️ **高风险五类字段**（收款账号 / 金额 / 利率 / 日期 / 证件号码）与**置信度 < 0.85** 的格子，"
          "`--apply-all` **不会**替你决定，必须逐个写编号。", ""]
    return "\n".join(L)


def render_fill_result(store: StoreV2, plan: dict, results: list[dict], cross: dict,
                       run_id: str, batch_no: str, *, signed: bool,
                       sign_kind: str | None = None, sign_actor: str | None = None) -> str:
    L = ["# 填报结果", "", "| 项 | 值 |", "| --- | --- |",
         "| 批次 | %s |" % batch_no, "| 运行号 | %s |" % run_id,
         "| 填成功 | **%d** 格 |" % sum(r["filled"] for r in results),
         "| 填失败 | %d 格 |" % sum(r["failed"] for r in results),
         "| 跨主体核对（本期口径） | 通过 %d，违规 %d，无法核对 %d |"
         % (cross["checked"], len(cross["violations"]), len(cross["unchecked"])),
         "| 签核 | %s |" % (
             "✅ **使用者已确认**（%s）" % sign_actor if signed else
             ("⚠️ **代理代填**（%s）——**不等于使用者检查过成品**，门禁四① **不认它**"
              % sign_actor if sign_kind == "declared-by-agent" else
              "🔴 **未签核 → 不许交付**")),
         "| **工具指纹** | `%s`（这套产物是**这个版本**的工具做的） |" % _tool_fp(),
         ""]
    L += ["## 逐份产物", "", "| 模板 | 产物 | 指纹 | 打开自检 | 格式说明 |",
          "| --- | --- | --- | --- | --- |"]
    for a in store.artifacts(run_id=run_id):
        L.append("| %s | %s | `%s` | %s | %s |" % (
            a.get("template_name") or "—", a["artifact_path"],
            (a["artifact_sha256"] or "")[:16],
            "✅" if a["opened_ok"] else ("🔴" if a["opened_ok"] == 0 else "—"),
            a["detail"] or ""))
    L += ["", "## 逐格结果", "", "| 字段 | 状态 | 值 | 说明 |", "| --- | --- | --- | --- |"]
    for r in store.conn.execute(
            "SELECT field,status,value,detail FROM fill_op WHERE run_id=? AND kind='field'"
            " ORDER BY id", (run_id,)):
        L.append("| `%s` | %s | %s | %s |" % (r[0], r[1], _cell(r[2]), r[3] or ""))
    opens = store.open_reviews()
    if opens:
        L += ["", "## 待你确认（%d 项）" % len(opens), "",
              "| # | 类别 | 字段/主体 | 原因 |", "| --- | --- | --- | --- |"]
        for r in opens:
            L.append("| %s | %s | %s | %s |" % (r["id"], r["kind"],
                                                r["field"] or r["label"], r["reason"]))
    L += ["", "## 交付门禁", "",
          "| # | 门禁 | 结果 |", "| --- | --- | --- |",
          "| 一 | 跨主体串数据 | %s |" % ("✅ 通过" if cross["ok"] else "🔴 违规"),
          "| 二 | 无中生有 | ✅ 通过（每个值都有来源类型） |",
          "| 三 | 格式未变 / 打得开 | 见上表「保真」列 |",
          "| 四 | 交付前置未满足 | %s |" % (
              "✅ 使用者已确认" if signed else
              ("⚠️ **代理代填，不算** —— 🔴 不许交付" if sign_kind == "declared-by-agent"
               else "🔴 未签核 —— **不许交付**")), ""]
    return "\n".join(L)


def _cell(value: Any) -> str:
    if value is None or str(value) == "":
        return "（空）"
    text = str(value).replace("|", "\\|")
    return text if len(text) <= 36 else text[:33] + "…"


# ==========================================================================
# 站点 ④ ⑭：查看 / 报告 / 追溯
# ==========================================================================
def cmd_db_show(args) -> Result:
    """④ 按键值对查看库里的内容（"整理和筛选"的界面）。"""
    res = Result("db-show")
    wr, store = _open(args)
    by = getattr(args, "sort_by", None) or "key"
    batch_no = getattr(args, "batch", None)
    entity_id = getattr(args, "entity_id", None)
    key = getattr(args, "key", None)
    min_conf = getattr(args, "min_confidence", None)

    sql = ("SELECT f.*, e.entity_type FROM fact f LEFT JOIN entity e ON e.id=f.entity_id"
           " WHERE f.superseded_by IS NULL")
    a: list = []
    if batch_no:
        sql += " AND f.batch_no=?"
        a.append(batch_no)
    if entity_id:
        sql += " AND f.entity_id=?"
        a.append(entity_id)
    if key:
        sql += " AND f.key LIKE ?"
        a.append("%" + key + "%")
    if min_conf is not None:
        sql += " AND COALESCE(f.confidence,1.0) >= ?"
        a.append(float(min_conf))
    rows = [dict(r) for r in store.conn.execute(sql + " ORDER BY f.id", tuple(a))]
    for r in rows:
        r["entity_name"] = store.entity_label(r["entity_id"])

    order = {"entity": lambda r: (r["entity_id"] or 0, r["key"]),
             "case": lambda r: (r["case_id"] or 0, r["key"]),
             "confidence": lambda r: (-(r["confidence"] if r["confidence"] is not None else 1.0),
                                      r["key"]),
             "batch": lambda r: (r["batch_no"] or "", r["key"]),
             "key": lambda r: (r["key"], r["entity_id"] or 0)}[by]
    rows.sort(key=order)

    out = _rdir(args, "db-show", batch_no or store.current_batch())
    L = ["# 库里的键值对 / db-show", "",
         "按 **%s** 排列；只列**当前有效**的行（被覆盖的旧行不算，它们仍在库里）。" % by, "",
         "| # | 字段 | 值 | 主体 | 批次 | 来源类型 | 出处 | 置信度 |",
         "| --- | --- | --- | --- | --- | --- | --- | --- |"]
    kind_label = {"source": "源文件", "computed": "算出来的", "user": "你确认的"}
    for i, r in enumerate(rows, start=1):
        L.append("| %d | `%s` | %s | %s | %s | %s | %s | %s |" % (
            i, r["key"], _cell(r["value"]), r["entity_name"], r["batch_no"] or "—",
            kind_label.get(r["source_kind"], r["source_kind"]), r["provenance"] or "—",
            "—" if r["confidence"] is None else "%.2f" % r["confidence"]))
    if not rows:
        L.append("| — | （没有匹配的行） | | | | | | |")
    p = write_text(out / "db_show.md", "\n".join(L) + "\n")
    res.add_artifact(p, "库内容清单（按 %s）" % by)
    res.data.update({"by": by, "rows": len(rows), "items": rows[:400],
                     "batches": sorted({r["batch_no"] for r in rows if r["batch_no"]}),
                     "summary": store.summary()})
    store.close()
    return res


def cmd_db_report(args) -> Result:
    """工作流报告：概览 / 缺失 / 待确认 / 冲突 / 识别键 / 审计链。"""
    res = Result("db-report")
    wr, store = _open(args)
    s = store.summary()
    out = _rdir(args, "db-report", store.current_batch())
    L = ["# 工作流报告 / Workflow Report", "", "## 概览", "", "| 项目 | 数量 |", "| --- | --- |"]
    for k, v in s.items():
        L.append("| %s | %s |" % (k, v))
    L.append("")

    gate = EK.gate(store.conn)
    L += ["## 可选识别资料", "",
          "| 项 | 值 |", "| --- | --- |",
          "| 主体总数 | %d |" % EK.summarize(store.conn)["entities"],
          "| 已有识别键 | %d |" % EK.summarize(store.conn)["with_key"],
          "| 未提供识别码或身份证号（不阻断） | %d |" % gate["missing_count"],
          "| 判定 | 可以继续处理 |", ""]
    if gate["missing"]:
        L += ["| # | 主体 | 缺什么 | 备注 |", "| --- | --- | --- | --- |"]
        for b in gate["missing"]:
            L.append("| %s | %s | %s | %s |" % (b["entity_id"], b["name"],
                                                "、".join(EK.KIND_LABEL[k] for k in b["missing"]),
                                                b["reason"] or "—"))
        L.append("")

    unassigned = store.unassigned_facts()
    L += ["## 待归属的键值对（红线一核对不了的那些）", "",
          "| 字段 | 值 | 批次 |", "| --- | --- | --- |"]
    for r in unassigned[:60]:
        L.append("| `%s` | %s | %s |" % (r["key"], _cell(r["value"]), r["batch_no"]))
    if not unassigned:
        L.append("| — | （全部已归属） | |")
    L.append("")

    facts = store.facts_by_key()
    missing = {k: v for k, v in facts.items() if not v.get("value")}
    if missing:
        L += ["## 源文件缺失的信息（%d）" % len(missing), "",
              "| 字段 | 出处 | 说明 |", "| --- | --- | --- |"]
        for k, v in missing.items():
            L.append("| `%s` | %s | %s |" % (k, v.get("provenance") or "",
                                             v.get("note") or "需人工填写"))
        L.append("")

    reviews = store.open_reviews()
    if reviews:
        L += ["## 待人工确认（%d）" % len(reviews), "",
              "| # | 类别 | 字段/主体 | 原因 | 建议 |", "| --- | --- | --- | --- | --- |"]
        for r in reviews:
            L.append("| %s | %s | %s | %s | %s |" % (
                r["id"], r["kind"], r["field"] or r["label"], r["reason"],
                _cell(r.get("candidates_json"))))
        L.append("")

    conflicts = store.conflicts(open_only=True)
    if conflicts:
        L += ["## 取值冲突（%d，**未覆盖**）" % len(conflicts), "",
              "要覆盖请跑 `db-merge` 出勾选清单——**这里不替你决定**。", "",
              "| # | 字段 | 库里现值 | 新来的值 | 主体 |", "| --- | --- | --- | --- | --- |"]
        for c in conflicts:
            L.append("| %s | `%s` | %s | %s | %s |" % (
                c["id"], c["key"], _cell(c["existing"]), _cell(c["incoming"]),
                store.entity_label(c["entity_id"])))
        L.append("")

    from .store_v2 import verify_chain

    ch = verify_chain(store.conn)
    # 校验这件事本身也要留痕（06 §2.1 的 `audit_verify_run`）。
    # ⚠️ 先算再写：写进去的这条事件不属于刚才校验的那段链。
    store.event("audit_verify_run", batch_no=store.current_batch() or "00000000-00",
                payload={"ok": ch["ok"], "checked": ch.get("checked"),
                         "problem": ch.get("problem"), "by": getattr(args, "by", "user")})

    L += ["## 工具指纹", "",
          "`%s` —— **这套记录是哪个版本的工具做的**。" % _tool_fp(),
          "改过工具（含 fork 出来的特殊版）指纹就会变；"\
          "分诊与复现都看它（[24 §4.4 C](24_歧义与模糊清单.md)）。", ""]
    L += ["## 审计链（06 §2.2）", "",
          "| 项 | 值 |", "| --- | --- |",
          "| 事件数 | %d |" % s["events"], "| 批次数 | %d |" % ch.get("batches", 0),
          "| 独立重算 | %s |" % ("✅ 通过" if ch["ok"] else "🔴 " + str(ch.get("problem"))),
          "| 算法 | SM3（本机不支持时退 SHA-256，事件长度自证） |", ""]

    p = write_text(out / "workflow_report.md", "\n".join(L) + "\n")
    res.add_artifact(p, "工作流报告（概览/识别键/待归属/缺失/待确认/冲突/审计链）")
    res.data.update({"summary": s, "identity_gate": gate, "unassigned": unassigned,
                     "open_reviews": reviews, "conflicts": conflicts, "chain": ch})
    store.close()
    return res


def cmd_db_review(args) -> Result:
    """裁定一条待确认项；裁定结果可以按**源侧四选**写回字段库。"""
    res = Result("db-review")
    wr, store = _open(args)
    batch_no = getattr(args, "batch", None) or store.current_batch() or allocate_batch(store.conn)
    if getattr(args, "list", False) or not getattr(args, "id", None):
        res.data["open_reviews"] = store.open_reviews()
        res.data["summary"] = store.summary()
        store.close()
        return res

    rid = int(args.id)
    row = store.conn.execute("SELECT * FROM review_queue WHERE id=?", (rid,)).fetchone()
    if row is None:
        store.close()
        raise OfficeKitError("没有这个待确认项：id=%s" % rid)
    answer = getattr(args, "answer", None)
    new_value = getattr(args, "new_value", None)
    if answer not in ("accept", "blank", "discard", "new"):
        store.close()
        raise OfficeKitError("裁定只能是 accept / new / discard / blank（源侧四选）：%r" % answer)
    store.resolve_review(rid, answer, answered_value=new_value,
                         decided_by=getattr(args, "by", None) or "user")

    field = getattr(args, "field", None) or row["field"]
    if answer == "accept" and field:
        # 接受建议值：值本来就在库里，**不需要改写**——只把裁定记下来。
        res.data["note"] = "接受了建议值——值本来就在库里，无需改写"
    elif answer == "new" and field:
        store.put_fact(batch_no, field, new_value, entity_id=row["entity_id"],
                       source_kind="user", provenance="你确认的",
                       note="来自待确认项 #%s" % rid, on_conflict="supersede")
        res.data["fact_updated"] = {"field": field, "value": new_value, "source_kind": "user"}
    elif answer == "discard" and field:
        store.discard_fact(batch_no, field, row["candidates_json"], entity_id=row["entity_id"],
                           reason="源侧复核：丢弃（待确认项 #%s）" % rid)
    res.data.update({"resolved": rid, "answer": answer, "field": field,
                     "summary": store.summary()})
    store.close()
    return res


def cmd_db_trace(args) -> Result:
    """追溯一个字段：出处、冲突、用在了哪些文档、经过哪些事件。"""
    res = Result("db-trace")
    wr, store = _open(args)
    res.data = store.trace(args.key)
    if not res.data["facts"]:
        res.warn("数据库里没有字段 %r" % args.key)
    store.close()
    return res


__all__ = [
    "AUTO_APPLY_THRESHOLD", "PROPOSE_THRESHOLD", "SYNONYMS", "score_candidate",
    "candidates_for", "extract_labels", "infer_roles", "build_fill_plan",
    "parse_fill_selection", "cmd_db_ingest", "cmd_db_roles", "cmd_db_propose",
    "cmd_db_rule", "cmd_db_fill", "cmd_db_report", "cmd_db_review", "cmd_db_trace",
    "cmd_db_show", "render_fill_plan",
]
