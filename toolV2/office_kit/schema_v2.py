r"""SCHEMA_VERSION 2 -- the 14 tables of phase 1.

GENERATED FROM THE DESIGN DOC.  Do not hand-edit the SQL below.

    source of truth : D:\DSH\project_bank_adjust\27_第一阶段详细设计.md  section 3.3
    ddl sha256      : c3127c4ced2771bced6aa4b53056e830f86ca4e91a0d20350a8c749697d5474b
    table count     : 15

`selftest_v2.py` re-extracts section 3.3 from the design doc and fails if it no
longer matches this constant, so the doc and the code cannot drift apart.
"""
from __future__ import annotations

SCHEMA_VERSION_V2 = 2
DDL_SHA256 = "c3127c4ced2771bced6aa4b53056e830f86ca4e91a0d20350a8c749697d5474b"

DDL_V2 = r'''-- 主体（企业或自然人）。D-1：隔离边界是主体，不是案子。
CREATE TABLE IF NOT EXISTS entity (
    id                        INTEGER PRIMARY KEY AUTOINCREMENT,
    entity_type               TEXT NOT NULL              -- 'legal_person' | 'natural_person'
                              CHECK (entity_type IN ('legal_person','natural_person')),
    uscc                      TEXT,                      -- 统一社会信用代码（法人）
    bank_no                   TEXT,                      -- 行内编号（与 uscc 一一对应）
    id_card                   TEXT,                      -- 身份证号（自然人）
    identity_incomplete       INTEGER NOT NULL DEFAULT 0,-- 识别键缺失=1（**待补录**，不是可停留状态，26 §1.1.1）
    identity_pending_reason   TEXT,
    note                      TEXT,
    created_at                TEXT NOT NULL,
    updated_at                TEXT NOT NULL
);
-- P-02：唯一索引必须含 entity_type，否则身份证号与信用代码理论上可能撞值
CREATE UNIQUE INDEX IF NOT EXISTS ux_entity_legal
    ON entity(uscc)         WHERE entity_type='legal_person'  AND uscc    IS NOT NULL;
CREATE UNIQUE INDEX IF NOT EXISTS ux_entity_legal_bankno
    ON entity(bank_no)      WHERE entity_type='legal_person'  AND bank_no IS NOT NULL;
CREATE UNIQUE INDEX IF NOT EXISTS ux_entity_natural
    ON entity(id_card)      WHERE entity_type='natural_person' AND id_card IS NOT NULL;

-- 主体历史名称 / 别名。D-1：一个主体可对应多个名字。
CREATE TABLE IF NOT EXISTS entity_name (
    id           INTEGER PRIMARY KEY AUTOINCREMENT,
    entity_id    INTEGER NOT NULL REFERENCES entity(id),
    name         TEXT NOT NULL,
    name_kind    TEXT NOT NULL DEFAULT '现用名' CHECK (name_kind IN ('现用名','曾用名','别名')),
    valid_from   TEXT,
    valid_to     TEXT,
    source_id    INTEGER REFERENCES source(id),
    legacy_session_id INTEGER,                     -- 迁移用：旧 session.id，保留以便回溯（§3.4）
    created_at   TEXT NOT NULL
);
CREATE INDEX IF NOT EXISTS ix_entity_name_entity ON entity_name(entity_id);
CREATE INDEX IF NOT EXISTS ix_entity_name_name   ON entity_name(name);

-- 主体基础信息（带时点版本）。P-01 + 26 §3.5：取"最新有效版"。
CREATE TABLE IF NOT EXISTS entity_profile (
    id           INTEGER PRIMARY KEY AUTOINCREMENT,
    entity_id    INTEGER NOT NULL REFERENCES entity(id),
    field        TEXT NOT NULL,                -- 标准字段名（见 18 字段册）
    value        TEXT,
    unit         TEXT,                         -- '元' | '万元' | NULL
    value_num    REAL,                         -- 归一到元（金额类才有）
    valid_from   TEXT,
    valid_to     TEXT,
    source_id    INTEGER REFERENCES source(id),
    created_at   TEXT NOT NULL
);
CREATE INDEX IF NOT EXISTS ix_entity_profile_lookup ON entity_profile(entity_id, field, valid_from);

-- 案子 = 一次授信批复（额度层）。26 §2.2。
-- ⚠️ 额度以**行内额度管理系统为准，我们只做对账**（18 §2.7.1）：因此**只存批复值**，
--    「已占用额度 / 可用额度」**不落库**（M3 等式是拿行内系统的数核对我们这一侧，26 §3.4）。
-- ⚠️ 同一主体**额度调整一次 = 新开一个案子**（D-1 原话"多个案子，按业务发生时间排序"），
--    所以本表不需要 valid_from/valid_to；"当日"概念只作用于**业务信息**（`fact.effective_*`，26 §3.5）。
CREATE TABLE IF NOT EXISTS "case" (
    id                      INTEGER PRIMARY KEY AUTOINCREMENT,
    batch_no                TEXT,                      -- 建档批次（YYYYMMDD-NN）
    entity_id               INTEGER NOT NULL REFERENCES entity(id),
    facility_no             TEXT,                      -- 授信批复编号
    credit_amount           REAL,                      -- 授信额度（元，批复值）
    credit_exposure_amount  REAL,                      -- 授信敞口额度（元，批复值）
    start_date              TEXT,                      -- 授信起始日
    end_date                TEXT,                      -- 授信到期日
    base_date               TEXT,                      -- 基准日 = 该笔业务的协议生效日期（26 §3.5）
    product_code            TEXT,                      -- 预留扩展口
    source_id               INTEGER REFERENCES source(id),
    legacy_session_id       INTEGER,                    -- 迁移用（§3.4）
    created_at              TEXT NOT NULL
);
CREATE INDEX IF NOT EXISTS ix_case_entity ON "case"(entity_id, start_date);

-- 主体在案子里的角色。26 §2.2：挂"主体 × 案子"；**可多身份**。
CREATE TABLE IF NOT EXISTS case_role (
    id           INTEGER PRIMARY KEY AUTOINCREMENT,
    entity_id    INTEGER NOT NULL REFERENCES entity(id),
    case_id      INTEGER NOT NULL REFERENCES "case"(id),
    role         TEXT NOT NULL
                 CHECK (role IN ('借款人','保证人','法定代表人')),   -- X-3：只做粗粒度，集合封闭
    evidence     TEXT,                                  -- 源文件里的原话/出处
    decided_by   TEXT NOT NULL DEFAULT 'human'
                 CHECK (decided_by IN ('source','human','auto')),  -- 'auto' = 源文件明写的路径（旧库就在用，见 §3.4.2）
    source_id    INTEGER REFERENCES source(id),
    legacy_session_id INTEGER,                          -- 迁移用（原 role_gate.session_id，§3.4）
    created_at   TEXT NOT NULL,
    UNIQUE(entity_id, case_id, role)                    -- ⚠️ 一条一笔一角色；同一主体可多行
);
CREATE INDEX IF NOT EXISTS ix_case_role_case ON case_role(case_id);
-- 源文件登记 + 副本 + 指纹
CREATE TABLE IF NOT EXISTS source (
    id            INTEGER PRIMARY KEY AUTOINCREMENT,
    batch_no      TEXT NOT NULL,                 -- YYYYMMDD-NN（26 §2.3）
    batch_label   TEXT,                          -- 人看的批次名（如"众森-2025-001"）；**不放进 batch_no**
    path          TEXT NOT NULL,                 -- 交付进来的原始路径
    copy_path     TEXT,                          -- in\<批次号>\ 下的副本（D-1：存整份副本）
    name          TEXT NOT NULL,
    sha256        TEXT NOT NULL,
    bytes         INTEGER NOT NULL,
    kind          TEXT,                          -- docx/xlsx/pdf/图片/...
    parse_status  TEXT NOT NULL DEFAULT 'pending'
                  CHECK (parse_status IN ('pending','ok','skipped','failed')),
    parse_note    TEXT,                          -- 跳过/失败原因（26 §9：不许静默跳过）
    legacy_session_id INTEGER,                   -- 迁移用（§3.4）
    ingested_at   TEXT NOT NULL,
    UNIQUE(path, sha256)
);
CREATE INDEX IF NOT EXISTS ix_source_batch ON source(batch_no);

-- 键值对。本项目最终目的的核心载体。
CREATE TABLE IF NOT EXISTS fact (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    batch_no        TEXT NOT NULL,
    source_id       INTEGER REFERENCES source(id),
    entity_id       INTEGER REFERENCES entity(id),
    case_id         INTEGER REFERENCES "case"(id),
    agreement_no    TEXT,                        -- 协议号（业务标识；单笔业务本期不单建表）
    key             TEXT NOT NULL,               -- 标准字段名
    value           TEXT,                        -- 原值（按文字存；协议号等保留前导 0）
    unit            TEXT,                        -- '元' | '万元'
    value_num       REAL,                        -- 归一到元的数值（仅金额类，26 §2.8）
    confidence      REAL,                        -- 提取置信度 0–1（26 §2.5）
    source_kind     TEXT NOT NULL DEFAULT 'source'
                    CHECK (source_kind IN ('source','computed','user')),  -- 红线二的三种可追溯来源
    provenance      TEXT,                        -- ⚠️ **人看的出处**，如 "0.1 第10行"、"人工确认"
                                                 --    迁移前老库把出处写在 origin 列，见 §3.4 步 6
    formula         TEXT,                        -- source_kind='computed' 时必填：算式
    inputs          TEXT,                        -- source_kind='computed' 时必填：输入值 JSON
    status          TEXT NOT NULL DEFAULT 'ok'
                    CHECK (status IN ('ok','missing','suspect')),
    superseded_by   INTEGER REFERENCES fact(id), -- 行只追加；旧行指向新行（17 §3.8）
    effective_from  TEXT,                        -- 版本时点（业务信息按 base_date 取）
    effective_to    TEXT,
    note            TEXT,
    legacy_session_id INTEGER,                   -- 迁移用（§3.4）
    created_at      TEXT NOT NULL
);
CREATE INDEX IF NOT EXISTS ix_fact_key    ON fact(key);
CREATE INDEX IF NOT EXISTS ix_fact_entity ON fact(entity_id, key);
CREATE INDEX IF NOT EXISTS ix_fact_agr    ON fact(agreement_no);
CREATE INDEX IF NOT EXISTS ix_fact_src    ON fact(source_id);

-- 冲突记录。26 §2.8：先归一后比较，归一后仍不等才记冲突。
CREATE TABLE IF NOT EXISTS fact_conflict (
    id            INTEGER PRIMARY KEY AUTOINCREMENT,
    batch_no      TEXT,
    key           TEXT NOT NULL,
    entity_id     INTEGER REFERENCES entity(id),
    agreement_no  TEXT,
    existing      TEXT,
    incoming      TEXT,
    existing_src  INTEGER REFERENCES source(id),
    incoming_src  INTEGER REFERENCES source(id),
    norm_rule     TEXT,                          -- 用哪条归一规则比的（金额→元 / 日期→YYYY-MM-DD / 文本）
    detected_at   TEXT NOT NULL,
    resolved      INTEGER NOT NULL DEFAULT 0,
    resolution    TEXT,                          -- 'took_existing' | 'took_incoming' | 'left_blank'
    resolved_by   TEXT,
    resolved_at   TEXT
);
CREATE INDEX IF NOT EXISTS ix_conflict_open ON fact_conflict(resolved, detected_at);
-- 字段契约。04 定义的类型/格式/必填/单位/枚举落到这里。
CREATE TABLE IF NOT EXISTS field_contract (
    id            INTEGER PRIMARY KEY AUTOINCREMENT,
    field         TEXT NOT NULL UNIQUE,          -- 标准字段名
    display_name  TEXT,                          -- 中文名（给使用者看）
    section       TEXT,                          -- 分组（企业基础信息 / 业务信息 / ...）
    dtype         TEXT,                          -- 'text' | 'amount' | 'date' | 'int' | 'enum'
    unit          TEXT,
    required      INTEGER NOT NULL DEFAULT 0,    -- **字段级默认**必填提示。
                                                 -- ⚠️ 判定"这个格子必不必填"**以 template_rule.is_required 为准**（26 §2.7）
    high_risk     INTEGER NOT NULL DEFAULT 0,    -- ⚠️ 1 = 高风险五类字段（26 §2.6）：
                                                 --    收款账号 · 金额 · 利率 · 日期 · 身份证号
                                                 --    置 1 的字段**无论置信度多高都必须过使用者复核**（26 §2.5）
    enum_json     TEXT,                          -- 封闭取值集合（业务类别本期为 NULL=自由文本）
    format_rule   TEXT,                          -- 'uscc18' | 'idcard18' | 'bankno19' | NULL（仅这三项）
    note          TEXT,
    created_at    TEXT NOT NULL,
    updated_at    TEXT NOT NULL
);

-- 模板登记
CREATE TABLE IF NOT EXISTS template (
    id             INTEGER PRIMARY KEY AUTOINCREMENT,
    batch_no       TEXT,
    path           TEXT NOT NULL,
    name           TEXT NOT NULL,
    sha256         TEXT NOT NULL,
    doc_kind       TEXT,                         -- 'docx' | 'xlsx'
    structure_ok   INTEGER,                      -- 结构自检：单一表头、无公式/宏/外链
    structure_note TEXT,
    legacy_session_id INTEGER,                   -- 迁移用（§3.4）
    registered_at  TEXT NOT NULL
);

-- 模板格子 × 字段 的映射规则（原表名 rule；template 是全局的，不按案子复制）
-- ⚠️ **本表的一行 = 目标表单上的一个格子**，所以：
--    · "应填字段"个数 = 该 template_id 下的行数（26 §2.7 完整率的分母）
--    · "必填字段" = is_required=1 的行（必填未满足 → 🟡 提示 + 入日志）
CREATE TABLE IF NOT EXISTS template_rule (
    id           INTEGER PRIMARY KEY AUTOINCREMENT,
    template_id  INTEGER NOT NULL REFERENCES template(id),
    field        TEXT NOT NULL,
    label        TEXT,                           -- 模板上的标签文字
    target_json  TEXT NOT NULL,                  -- 位置（1 基：表/行/列 或 段落/run 锚点）
    is_required  INTEGER NOT NULL DEFAULT 0,     -- 该模板该格是否必填（26 §2.7）
    match_kind   TEXT NOT NULL DEFAULT 'exact'
                 CHECK (match_kind IN ('exact','normalized')),
    confidence   REAL,                           -- 匹配置信度 0–1
    decided_by   TEXT CHECK (decided_by IN ('human','model','auto')),  -- 与 CLI --decided-by 取值一致
    legacy_session_id INTEGER,                   -- 迁移用（原 rule 无此列，留空）
    created_at   TEXT NOT NULL,
    UNIQUE(template_id, field)
);
-- 删除会让旧编号指向空处，历史填报记录也会失去依据；废弃规则只做停用。
CREATE TRIGGER IF NOT EXISTS template_rule_no_delete
BEFORE DELETE ON template_rule
BEGIN
    SELECT RAISE(ABORT, 'template_rule 不允许按编号删除；请按模板和字段停用');
END;
CREATE TABLE IF NOT EXISTS template_rule_disabled (
    template_id INTEGER NOT NULL,
    field TEXT NOT NULL,
    disabled_at TEXT NOT NULL,
    reason TEXT NOT NULL,
    PRIMARY KEY (template_id, field)
);
-- 填充记录。一张表三种行，靠 kind 区分：
--   kind='run'      每次运行一行：承载【执行前快照】与本次运行的状态
--   kind='artifact' **每一份产物一行**：承载【产物指纹 / 打开自检 / 签核结果】
--   kind='field'    每个字段一行：承载逐字段结果
--
-- ⚠️ 为什么产物与签核必须单独成行（不能挤在 run 行里）：
--    一次运行**可以跨多个主体、多个案子、产出多份目标文档**（26 §2.2），
--    而签核是**一批一次 + 每份一次**（N-3），§7.1 红线四还要**逐份**查 signoff_decision。
--    一份产物一行，这三件事才都表达得出来。
CREATE TABLE IF NOT EXISTS fill_op (
    id                  INTEGER PRIMARY KEY AUTOINCREMENT,
    run_id              TEXT NOT NULL,            -- 一次运行（如 批次号-序号）
    batch_no            TEXT NOT NULL,
    kind                TEXT NOT NULL CHECK (kind IN ('run','artifact','field')),
    entity_id           INTEGER REFERENCES entity(id),
    template_id         INTEGER REFERENCES template(id),
    template_rule_id    INTEGER REFERENCES template_rule(id),
    fact_id             INTEGER REFERENCES fact(id),
    field               TEXT,
    value               TEXT,
    status              TEXT NOT NULL             -- 见下方"status 取值"
                        CHECK (status IN ('running','ok','missing','blocked','skipped','aborted','rejected')),
    detail              TEXT,
    -- 以下仅 kind='run' 填写
    preflight_json      TEXT,                     -- 执行前快照（模板哈希 + 目标位置清单）
    -- 以下仅 kind='artifact' 填写（**一份产物一行**）
    artifact_path       TEXT,                     -- out\<批次号>\<主体>\ 下的产物
    artifact_sha256     TEXT,                     -- 产物指纹（N-1：记这里，不单建表）
    opened_ok           INTEGER,                  -- 产物可打开自检
    signoff_scope       TEXT CHECK (signoff_scope IN ('batch','document')),
    signoff_decision    TEXT CHECK (signoff_decision IN ('approve','reject')),
    signoff_actor       TEXT,                     -- 签核人 = 使用者本人（26 §3.8）
    signoff_at          TEXT,
    signoff_fingerprint TEXT,                     -- 26 §3.7：快照 + 产物 的指纹
    legacy_session_id   INTEGER,                  -- 迁移用（§3.4）
    created_at          TEXT NOT NULL
);
CREATE INDEX IF NOT EXISTS ix_fill_op_run  ON fill_op(run_id);
CREATE INDEX IF NOT EXISTS ix_fill_op_batch ON fill_op(batch_no);
CREATE INDEX IF NOT EXISTS ix_fill_op_tpl  ON fill_op(template_id, kind);
CREATE INDEX IF NOT EXISTS ix_fill_op_art  ON fill_op(run_id, entity_id, template_id) WHERE kind='artifact';
-- 待确认项（原 review_item + role_question 合并）
CREATE TABLE IF NOT EXISTS review_queue (
    id            INTEGER PRIMARY KEY AUTOINCREMENT,
    batch_no      TEXT,
    run_id        TEXT,
    kind          TEXT NOT NULL CHECK (kind IN ('field','role','conflict')),
    entity_id     INTEGER REFERENCES entity(id),
    case_id       INTEGER REFERENCES "case"(id),
    template_id   INTEGER REFERENCES template(id),
    field         TEXT,
    label         TEXT,
    reason        TEXT NOT NULL,                 -- 为什么要问你（低置信度 / 高风险 / 角色未明 / 冲突）
    candidates_json TEXT,                        -- 建议值 + 候选
    answer        TEXT,                          -- 'accept' | 新值 | 'blank'
    answered_value TEXT,                         -- 选择"输入新值"时的值
    decided_by    TEXT,
    legacy_session_id INTEGER,                   -- 迁移用（原 review_item / role_question 的 session_id）
    resolved_at   TEXT,
    created_at    TEXT NOT NULL
);
CREATE INDEX IF NOT EXISTS ix_review_open ON review_queue(resolved_at, created_at);

-- 操作日志 / 审计事件链（哈希链）：一条链，按 batch 成链
-- ⚠️ 列名与 [06 §2.2](06_审计日志与交叉验证.md) 的必备字段**逐字对齐**——哈希公式用的是这些列。
CREATE TABLE IF NOT EXISTS event (
    id            INTEGER PRIMARY KEY AUTOINCREMENT,
    batch_no      TEXT NOT NULL,
    run_id        TEXT,
    seq           INTEGER NOT NULL,              -- **批次内**序号，从 1 连续（链按 batch 成链）
    occurred_at   TEXT NOT NULL,                 -- UTC 时间 + 时区（06 §2.2；G-B8）
    actor_id      TEXT,                          -- 谁（使用者账号 / 'agent'）
    actor_name    TEXT,
    event_type    TEXT NOT NULL,                 -- 见 06 §2.1 唯一权威枚举
    case_id       INTEGER,                       -- 06 §2.2 的必备字段之一（见下方说明）
    target        TEXT,                          -- 作用于什么
    entity_id     INTEGER,                       -- 纳入哈希（17 §3.9）
    payload_json  TEXT,                          -- 载荷（含明文值，D-3 允许）
    payload_hash  TEXT NOT NULL,                 -- 载荷哈希
    value_text    TEXT,                          -- 明文值（D-3：哈希 + 明文都存）
    prev_hash     TEXT,                          -- 上一条的 hash
    hash          TEXT NOT NULL,                 -- H(prev_hash‖event_type‖payload_hash‖occurred_at‖actor_id‖run_id‖entity_id‖seq)
                                                 -- 公式以 06 §2.2 为准；默认 SM3（D-5）
    session_ref   TEXT,                          -- 关联 DSH 会话记录（能力②的第一层）
    UNIQUE(batch_no, seq)
);
CREATE INDEX IF NOT EXISTS ix_event_batch ON event(batch_no, seq);
'''

# ---------------------------------------------------------------- value maps
# Every map below is copied from doc 27 section 3.4.2 / 3.4.4 / 3.4.5 (v0.4).

#: old `fact.origin` -> (new `fact.source_kind`, new `fact.provenance`)
def map_origin(origin):
    if origin is None or origin == "":
        return ("source", None)
    if origin in ("\u4eba\u5de5\u786e\u8ba4", "user"):
        return ("user", origin)
    if origin == "computed":
        return ("computed", origin)
    return ("source", origin)   # a provenance string: keep it verbatim


#: old `fill_op.status` -> new `fill_op.status`  (\u26a0 `filled` is the live success value)
STATUS_MAP = {
    "filled": "ok",
    "ok": "ok",
    "missing": "missing",
    "blocked": "blocked",
    "skipped": "skipped",
    "aborted": "aborted",
}


#: old `role_gate.role` -> new `case_role.role`, key-word rule, first hit wins.
#: The live vocabulary includes composite forms such as \u501f\u6b3e\u4eba\u6cd5\u5b9a\u4ee3\u8868\u4eba,
#: so a per-value table would miss them (doc 27 section 3.4.2).
def map_role(role):
    r = role or ""
    if "\u6cd5\u5b9a\u4ee3\u8868\u4eba" in r:
        return "\u6cd5\u5b9a\u4ee3\u8868\u4eba"
    if "\u4fdd\u8bc1\u4eba" in r:
        return "\u4fdd\u8bc1\u4eba"
    if "\u501f\u6b3e\u4eba" in r:
        return "\u501f\u6b3e\u4eba"
    return None


#: fact keys whose *value* is an entity name.  Taken from harness.infer_roles'
#: entity_keys -- these are the keys live writers actually produce.
SUBJECT_NAME_KEYS = [
    "\u501f\u6b3e\u4eba\u540d\u79f0",
    "\u4fdd\u8bc1\u4eba\u540d\u79f0",
    "\u501f\u6b3e\u4eba\u6cd5\u5b9a\u4ee3\u8868\u4eba",
    "\u4fdd\u8bc1\u4eba\u6cd5\u5b9a\u4ee3\u8868\u4eba",
    "\u4e2a\u4eba\u4fdd\u8bc1\u4eba\u59d3\u540d",
]


def entity_type_for(name, *, personal_keys=(), roles=()):
    """doc 27 section 3.4.2: never guess from the name.

    Default is legal_person (almost every subject is a company).  Only an
    explicit natural-person signal -- the name came from a personal-name key,
    or the subject is recorded as a personal guarantor -- promotes it.
    Returns (entity_type, needs_confirmation).
    """
    if name in personal_keys or "\u4e2a\u4eba\u4fdd\u8bc1\u4eba" in (roles or ()):
        return ("natural_person", False)
    return ("legal_person", True)
