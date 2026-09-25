"""Command-line interface: python -m office_kit <command> [options]."""
from __future__ import annotations

import argparse
import os
import sys
import traceback
import warnings
from pathlib import Path

from . import __version__
from .common import OfficeKitError, Result

# Third-party import noise (docxcompose/rapidocr touch the deprecated
# pkg_resources API) must never reach stdout, which carries only the JSON envelope.
warnings.filterwarnings("ignore", message=r".*pkg_resources.*")
warnings.filterwarnings("ignore", category=DeprecationWarning, module=r".*pkg_resources.*")


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="office_kit",
        description="Local office toolkit: data cleaning/analysis, document parsing & generation, OCR, file organization.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="Every command prints one JSON envelope. Reports are written to files under --out.",
    )
    p.add_argument("--version", action="version", version=f"office_kit {__version__}")
    sub = p.add_subparsers(dest="command", required=True, metavar="COMMAND")

    def add_common(sp, *, out_default: bool = True):
        sp.add_argument("--out", help="directory for produced artifacts (default: ./_office_out/<command>)")
        return sp

    # ---------------------------------------------------------------- data
    sp = sub.add_parser("inspect", help="peek at files: encoding, sheets, shape, preview, obvious problems")
    sp.add_argument("input", help="file, directory or glob")
    sp.add_argument("--sheet", help="worksheet name for Excel inputs")
    sp.add_argument("--header", type=int, default=0, help="header row index (default 0)")
    sp.add_argument("--preview-rows", type=int, default=5)
    add_common(sp)

    sp = sub.add_parser("clean", help="clean a messy table (names, types, nulls, dupes, outliers)")
    sp.add_argument("input")
    sp.add_argument("--out", help="artifact directory")
    sp.add_argument("--sheet")
    sp.add_argument("--header", type=int, default=0)
    sp.add_argument("--dedupe", dest="dedupe", action="store_true", default=True)
    sp.add_argument("--no-dedupe", dest="dedupe", action="store_false")
    sp.add_argument("--dedupe-keys", help="comma-separated subset of columns for duplicate detection")
    sp.add_argument("--outliers", choices=["flag", "clip", "none"], default="flag",
                    help="flag adds <col>_is_outlier columns; clip caps them at the IQR fence")
    sp.add_argument("--no-fullwidth", dest="normalize_fullwidth", action="store_false", default=True)
    sp.add_argument("--no-fuzzy-categories", dest="fuzzy_categories", action="store_false", default=True)
    sp.add_argument("--keep-column-names", action="store_true", default=False)
    sp.add_argument("--drop-empty", dest="drop_empty", action="store_true", default=True)
    sp.add_argument("--no-drop-empty", dest="drop_empty", action="store_false")
    sp.add_argument("--drop-unnamed", dest="drop_unnamed", action="store_true", default=True)
    sp.add_argument("--rules", help="JSON file with rename_columns / drop_columns")
    sp.add_argument("--emit-csv", action="store_true", default=False)

    sp = sub.add_parser("profile", help="deep per-column statistics, issues and correlations")
    sp.add_argument("input")
    sp.add_argument("--out")

    sp.add_argument("--sheet")
    sp.add_argument("--header", type=int, default=0)
    sp.add_argument("--top-n", type=int, default=10, help="how many top values per text column")
    sp.add_argument("--emit-markdown", dest="emit_markdown", action="store_true", default=True)
    sp.add_argument("--no-markdown", dest="emit_markdown", action="store_false")
    sp.add_argument("--emit-json", action="store_true", default=False)

    sp = sub.add_parser("pivot", help="group/pivot aggregation, exported to Excel")
    sp.add_argument("input")
    sp.add_argument("--out")
    sp.add_argument("--sheet")
    sp.add_argument("--header", type=int, default=0)
    sp.add_argument("--index", help="row grouping columns, comma separated")
    sp.add_argument("--columns", help="column grouping field(s)")
    sp.add_argument("--values", help="value columns to aggregate")
    sp.add_argument("--agg", help="e.g. 'amount:sum,qty:mean' or just 'sum' / '平均'")
    sp.add_argument("--name", help="output basename")
    sp.add_argument("--sort-by", help="sort the result by this column")
    sp.add_argument("--desc", action="store_true", default=False)

    sp = sub.add_parser("chart", help="make a chart image (bar/line/pie/scatter/hist/box/heatmap)")
    sp.add_argument("input")
    sp.add_argument("--out")
    sp.add_argument("--sheet")
    sp.add_argument("--header", type=int, default=0)
    sp.add_argument("--type", default="auto", help="bar, barh, line, area, pie, scatter, hist, box, heatmap, auto")
    sp.add_argument("--x", help="x / category column")
    sp.add_argument("--y", help="value column(s), comma separated")
    sp.add_argument("--group", help="series grouping column")
    sp.add_argument("--agg", default="sum", help="sum/mean/count/min/max/median/none")
    sp.add_argument("--top", type=int, default=30, help="keep only the top N categories")
    sp.add_argument("--bins", type=int, default=20, help="histogram bins")
    sp.add_argument("--title")
    sp.add_argument("--name", help="output basename")
    sp.add_argument("--figsize", help="WxH in inches, e.g. 12x7")
    sp.add_argument("--rotate", type=int, help="x tick rotation")

    sp = sub.add_parser("report", help="full analysis report: Excel with charts + markdown")
    sp.add_argument("input")
    sp.add_argument("--out")
    sp.add_argument("--sheet")
    sp.add_argument("--header", type=int, default=0)
    sp.add_argument("--title")
    sp.add_argument("--name", help="output basename")
    sp.add_argument("--group", help="dimension to break down by")
    sp.add_argument("--metrics", help="numeric columns to aggregate")
    sp.add_argument("--top-n", type=int, default=10)
    sp.add_argument("--emit-markdown", dest="emit_markdown", action="store_true", default=True)
    sp.add_argument("--no-markdown", dest="emit_markdown", action="store_false")

    sp = sub.add_parser("merge", help="combine many tables into one")
    sp.add_argument("input", nargs="+", help="files / directories / globs")
    sp.add_argument("--out")
    sp.add_argument("--sheet")
    sp.add_argument("--header", type=int, default=0)
    sp.add_argument("--mode", default="auto", choices=["auto", "concat", "align", "join"])
    sp.add_argument("--key", help="join key columns (mode=join)")
    sp.add_argument("--how", default="left", help="join how: left/right/inner/outer")
    sp.add_argument("--dedupe", action="store_true", default=False)
    sp.add_argument("--dedupe-keys")
    sp.add_argument("--no-source-column", dest="add_source", action="store_false", default=True)
    sp.add_argument("--name")

    sp = sub.add_parser("compare", help="diff two tables by key")
    sp.add_argument("input", nargs="+")
    sp.add_argument("--out")
    sp.add_argument("--sheet")
    sp.add_argument("--key", help="key columns; auto-detected when omitted")

    sp = sub.add_parser("data-convert", help="convert table formats (csv/xlsx/json/parquet/md)")
    sp.add_argument("input", nargs="+")
    sp.add_argument("--out")
    sp.add_argument("--sheet")
    sp.add_argument("--header", type=int, default=0)
    sp.add_argument("--to", required=True, help="xlsx, csv, json, md, parquet, tsv")
    sp.add_argument("--combine", action="store_true", help="also write one markdown file with all tables")

    # ------------------------------------------------------------ documents
    sp = sub.add_parser("extract", help="read PDF/Word/PPT/Excel -> text, markdown, tables, JSON")
    sp.add_argument("input", nargs="+")
    sp.add_argument("--out")
    sp.add_argument("--format", default="markdown",
                    help="markdown | text | json | csv | all (default markdown)")
    sp.add_argument("--password", help="password for encrypted PDFs")
    sp.add_argument("--images", action="store_true", help="also extract embedded images")
    sp.add_argument("--preview-rows", type=int, default=20)

    sp = sub.add_parser("build", help="create a Word/PowerPoint/Excel/PDF/HTML document")
    sp.add_argument("--kind", required=True, choices=["docx", "pptx", "xlsx", "pdf", "html"])
    sp.add_argument("--spec", help="JSON spec file (blocks/sections/slides/sheets)")
    sp.add_argument("--markdown", help="markdown source file")
    sp.add_argument("--text", help="plain text source file")
    sp.add_argument("--out")
    sp.add_argument("--name", help="output basename")
    sp.add_argument("--title")
    sp.add_argument("--via", choices=["auto", "matplotlib"], default="auto", help="PDF engine")
    sp.add_argument("--no-title", dest="with_title", action="store_false", default=True)

    sp = sub.add_parser("doc-convert", help="convert documents; uses MS Office for high fidelity")
    sp.add_argument("input", nargs="+")
    sp.add_argument("--to", required=True, help="pdf, docx, xlsx, pptx, csv, html, md, png")
    sp.add_argument("--out")
    sp.add_argument("--engine", choices=["auto", "office", "python"], default="auto",
                    help="office = require native Office fidelity and fail otherwise")

    sp = sub.add_parser("pdf", help="PDF operations: merge, split, pages, rotate, encrypt, info, images")
    sp.add_argument("op", choices=["merge", "split", "extract-pages", "rotate", "encrypt",
                                   "info", "images", "to-images"])
    sp.add_argument("input", nargs="+")
    sp.add_argument("--out")
    sp.add_argument("--pages", help="page ranges, e.g. 1-3,5,8")
    sp.add_argument("--every", type=int, default=1, help="split: pages per output file")
    sp.add_argument("--angle", type=int, default=90, help="rotate: degrees")
    sp.add_argument("--password", help="password for an encrypted input")
    sp.add_argument("--new-password", help="encrypt: password to set")
    sp.add_argument("--scale", type=float, default=2.0, help="to-images: render scale")
    sp.add_argument("--image-format", default="png", choices=["png", "jpg", "jpeg"])
    sp.add_argument("--name")

    sp = sub.add_parser("fill", help="fill a docx/xlsx/pptx template with JSON data")
    sp.add_argument("--template", required=True)
    sp.add_argument("--data", required=True, help="JSON file: object, or list with --repeat")
    sp.add_argument("--repeat", action="store_true", help="render one output per record")
    sp.add_argument("--out")
    sp.add_argument("--name")

    sp = sub.add_parser(
        "fillmap",
        help="表单批量填充：按字段字典+模板映射填表，保留原格式，缺失留空并出人工确认清单",
    )
    sp.add_argument("input", nargs="+", help="模板文件（.docx）")
    sp.add_argument("--profile", required=True, help="字段字典 JSON（值的唯一来源）")
    sp.add_argument("--mapping", required=True, help="模板映射 JSON（标签/单元格 -> 字段）")
    sp.add_argument("--out")
    sp.add_argument("--strict", dest="strict", action="store_true", default=False,
                    help="没有任何输出时直接报错")
    sp.add_argument("--no-strict", dest="strict", action="store_false")
    sp.add_argument("--engine", choices=["xml", "pydocx"], default="xml",
                    help="xml=只重写 document.xml，其余文件逐字节不变（保真最高）；"
                         "pydocx=用 python-docx 重写，会刷新包内其他部件")

    sp = sub.add_parser("to-docx", help="批量把旧版 .doc/.xls/.ppt 转成新格式（走本机 Office）")
    sp.add_argument("input", nargs="+")
    sp.add_argument("--out")

    # ------------------------------------------------- 持久化工作流 workflow
    db_help = "SQLite 工作流数据库（SCHEMA v2）。**不给就自动用工作区里的 db\\\\workflow.db**"

    sp = sub.add_parser(
        "work",
        help="工作区：指定一个文件夹，库/源文件副本/产物/日志/临时区全都长在它里面")
    sp.add_argument("--init", action="store_true",
                    help="把这个文件夹变成工作区（建 in/ out/ db/ logs/ work/ + 写标记文件）")
    sp.add_argument("--status", action="store_true", help="看看现在认领的是哪个工作区（默认）")
    sp.add_argument("--suggest", action="store_true",
                    help="工作区还没定时，给使用者一个**具体建议**（目录 + 文件夹名 + 理由）")
    sp.add_argument("--brief", action="store_true",
                    help="打印「动手前该对使用者说的三行」（工具 / 工作区 / 权限），照着念")
    sp.add_argument("--triage", action="store_true",
                    help="临时编程分诊：共性问题（工具升级）还是特殊问题（复制到本会话工作区）")
    sp.add_argument("--for-what", dest="for_what",
                    help="配 --suggest：这次要干的活儿/材料名，用来起文件夹名")
    sp.add_argument("--kind", choices=["common", "special"],
                    help="配 --triage：common=共性问题（工具升级）/ special=特殊问题（本会话特殊版）")
    sp.add_argument("--reason", help="配 --triage：为什么这么判（必填，要给人看）")
    sp.add_argument("--fork-name", dest="fork_name",
                    help="配 --triage --kind special：副本的名字（默认「特殊版」）")
    sp.add_argument("--by", default="user",
                    help="配 --triage：谁批的（写进 FORK.md 与事件链）")
    sp.add_argument("--at", help="工作区文件夹（配 --init；不给就用当前目录）")
    sp.add_argument("--work", dest="work", help="要查看的工作区（配 --status）")
    sp.add_argument("--note", help="给这个工作区写一句备注（记在标记文件里）")
    sp.add_argument("--json", action="store_true", help="原样打印 JSON，不渲染报告")
    sp.add_argument("--out", help="把工作区报告另存一份到这里")

    def add_work(sp):
        """所有跑数据库的命令都挂上工作区选项。"""
        sp.add_argument("--work", dest="work",
                        help="工作区文件夹（库/副本/产物/日志都在这下面）")
        sp.add_argument("--db", help=db_help)
        return sp

    sp = add_work(sub.add_parser("db-rules-export", help="只保存模板填写位置，不保存客户事实"))
    sp.add_argument("input", nargs="+", help="已在本工作区设好规则的目标模板")
    sp.add_argument("--out", required=True, help="规则包 JSON 的保存路径")

    sp = add_work(sub.add_parser("db-rules-import", help="相同模板直接复用位置规则"))
    sp.add_argument("pack", help="之前导出的规则包 JSON")
    sp.add_argument("input", nargs="+", help="本次要填写的目标模板")
    sp.add_argument("--batch", help="本次批次号")

    sp = add_work(sub.add_parser("db-ingest", help="把字段字典/信息源导入工作流数据库（登记源文件+抽键值对+建主体+认角色）"))
    # --db 已由 add_work 挂上（可选：不给就用工作区）
    sp.add_argument("--profile", required=True, help="字段字典 JSON（可含 source 出处）")
    sp.add_argument("--source", nargs="*", help="信息源文件（登记 + 存副本，并用来判角色）")
    sp.add_argument("--batch", help="批次号 YYYYMMDD-NN（默认按当天自动分配）")
    sp.add_argument("--label", help="人看的批次名（如 众森-2026-001），不放进批次号")
    sp.add_argument("--copy-root", dest="copy_root",
                    help="把源文件整份副本存到 <这里>\\<批次号>\\（D-1 要求存副本）")
    sp.add_argument("--default-entity", dest="default_entity",
                    help="这批材料的默认主体（名字或 id）；不带角色前缀的键都归它")
    sp.add_argument("--by", default="user")

    sp = add_work(sub.add_parser("db-rule-disable", help="按模板和字段停用错误规则，保留旧记录"))
    sp.add_argument("--template", required=True, help="目标模板完整路径")
    sp.add_argument("--field", required=True, help="当前规则的字段名")
    sp.add_argument("--expected-target", required=True,
                    help="当前规则的位置 JSON；位置变了就拒绝，防止依据旧记录清理")
    sp.add_argument("--reason", required=True, help="停用原因")

    sp = add_work(sub.add_parser(
        "db-absorb",
        help="从材料里读出候选键值对，摊成一张**人能审的中文表**（不写库；审完再 db-ingest）"))
    sp.add_argument("input", nargs="+", help="材料（.docx/.xlsx/.md/.txt/.csv）")
    sp.add_argument("--batch")
    sp.add_argument("--out", help=r"草稿放哪（默认 <工作区>\out\<批次>\_报告\db-absorb\）")
    sp.add_argument("--default-entity", dest="default_entity",
                    help="认不出归属的那些键，默认算谁的")
    sp.add_argument("--assign", action="append",
                    help="第几条归谁：编号=主体名，如 --assign 4=广东众森实业发展有限公司")
    sp.add_argument("--set", action="append",
                    help="第几条改成别的值：编号=值，如 --set 2=顾红军")
    sp.add_argument("--add", action="append",
                    help="材料里没有、由使用者补一条：键=值，如 --add 用信用途=采购原材料"
                         "（记 source_kind='user'）")
    sp.add_argument("--drop", action="append",
                    help="第几条不要：编号（不入库）")
    sp.add_argument("--by", default="user")

    sp = add_work(sub.add_parser("db-roles", help="人物/主体角色确认（🟡 提示级，不阻断填报）"))
    # --db 已由 add_work 挂上（可选：不给就用工作区）
    sp.add_argument("--id", type=int, help="要回答的待确认项 id")
    sp.add_argument("--answer", help="该主体的角色：借款人 / 保证人 / 法定代表人")
    sp.add_argument("--entity", help="主体名称（配合 --role 直接指定，无需 id）")
    sp.add_argument("--role", help="角色，如 借款人/保证人/法定代表人")
    sp.add_argument("--batch", help="批次号（默认取库里最新）")
    sp.add_argument("--by", default="user")

    sp = add_work(sub.add_parser("db-propose", help="按模板标签生成候选规则并给出置信度"))
    sp.add_argument("input", nargs="+", help="目标模板")
    # --db 已由 add_work 挂上（可选：不给就用工作区）
    sp.add_argument("--batch")
    sp.add_argument("--by", default="user")
    sp.add_argument("--out")

    sp = add_work(sub.add_parser("db-rule", help="写入/覆盖一条规则（人工或大模型裁决结果）"))
    # --db 已由 add_work 挂上（可选：不给就用工作区）
    sp.add_argument("--template", required=True)
    sp.add_argument("--field", required=True, help="对应的字段键")
    sp.add_argument("--label", help="模板中的标签文字（用于报告）")
    sp.add_argument("--anchor", help="标签，如 '客户名称：'")
    sp.add_argument("--before", help="右侧闭合标签（可空）")
    sp.add_argument("--target", help="完整 target JSON（单元格等复杂规则用）")
    sp.add_argument("--required", action="store_true",
                    help="这个格子是必填（is_required=1；完整率与提示都看它）")
    sp.add_argument("--match-kind", dest="match_kind", default="exact")
    sp.add_argument("--confidence", type=float)
    sp.add_argument("--decided-by", dest="decided_by", default="human",
                    help="human / model / auto")
    sp.add_argument("--batch")
    sp.add_argument("--by", default="user")

    sp = add_work(sub.add_parser(
        "db-fill",
        help="预演 → 校验 → 保真填充 → 签核（识别码和证件号码可选）"))
    sp.add_argument("input", nargs="+", help="目标模板（.docx / .xlsx）")
    # --db 已由 add_work 挂上（可选：不给就用工作区）
    sp.add_argument("--out")
    sp.add_argument("--batch", help="批次号（默认取库里最新）")
    sp.add_argument("--engine", choices=["xml", "pydocx"], default="xml",
                    help="xml=只重写 document.xml，其余部件逐字节不变（默认，保真可证明）；"
                         "pydocx=用 python-docx 重写，会重新序列化 run 属性")
    sp.add_argument("--plan", help="预演清单；默认从办公数据库读取确认后执行，缺项时转交原生弹窗")
    sp.add_argument("--confirmed", action="store_true",
                    help="从本工作区数据库读取已确认的选择，一次填完本计划中的目标文件")
    sp.add_argument("--select", help="目标侧三选之一：接受建议，如 1,3")
    sp.add_argument("--new", action="append",
                    help="目标侧三选之二：输入新值，如 --new 2=广东众森实业发展有限公司")
    sp.add_argument("--blank", action="append", help="目标侧三选之三：留空，如 --blank 4")
    sp.add_argument("--use", action="append",
                    help="这一格用**哪个主体**的值：编号=主体名或#编号，如 --use 3=广东众森实业发展有限公司"
                         "（只在「同一键属于多个主体」时才需要）")
    sp.add_argument("--apply-all", dest="apply_all", action="store_true",
                    help="全部自动填的那批（**不含**高风险五类与置信度不足的格子）")
    sp.add_argument("--sign",
                    help="⚠️ 记「**代理代填**」的签核声明：填姓名即记 approve，"
                         "但**不等于使用者检查过成品**，交付门禁四① 不认它")
    sp.add_argument("--sign-confirmed", dest="sign_confirmed",
                    help="使用者**本人在对话里确认过**：填他的姓名，且**必须**同时给 --quote 原话")
    sp.add_argument("--quote", help="配 --sign-confirmed：使用者的**原话**（可拿去 DSH 会话记录对照）")
    sp.add_argument("--by", default="user")

    sp = add_work(sub.add_parser("db-show", help="按键值对查看库里的内容（按主体/案子/置信度/批次/键）"))
    # --db 已由 add_work 挂上（可选：不给就用工作区）
    sp.add_argument("--sort-by", dest="sort_by", default="key",
                    choices=["key", "entity", "case", "confidence", "batch"])
    sp.add_argument("--entity-id", dest="entity_id", type=int)
    sp.add_argument("--batch")
    sp.add_argument("--key", help="字段名（支持部分匹配）")
    sp.add_argument("--min-confidence", dest="min_confidence", type=float)
    sp.add_argument("--actor", default="user")
    sp.add_argument("--out")

    sp = add_work(sub.add_parser("db-report", help="工作流报告：概览/识别键/待归属/缺失/待确认/冲突/审计链"))
    # --db 已由 add_work 挂上（可选：不给就用工作区）
    sp.add_argument("--by", default="user")
    sp.add_argument("--out")

    sp = add_work(sub.add_parser("db-review", help="查看或裁定待确认项（源侧四选；裁定后写回字段库）"))
    # --db 已由 add_work 挂上（可选：不给就用工作区）
    sp.add_argument("--list", action="store_true", help="列出未确认项")
    sp.add_argument("--id", type=int, help="要裁定的待确认项 id")
    sp.add_argument("--answer", choices=["accept", "new", "discard", "blank"],
                    help="accept=接受建议 / new=改后入库 / discard=丢弃 / blank=留空待查")
    sp.add_argument("--new-value", dest="new_value", help="answer=new 时你输入的值")
    sp.add_argument("--field", help="裁定结果写回哪个字段键（默认取待确认项自己的字段）")
    sp.add_argument("--batch")
    sp.add_argument("--by", default="user")

    sp = add_work(sub.add_parser("db-trace", help="追溯一个字段：出处、冲突、用在了哪些文档"))
    # --db 已由 add_work 挂上（可选：不给就用工作区）
    sp.add_argument("--key", required=True)
    sp.add_argument("--by", default="user")

    sp = add_work(sub.add_parser(
        "db-entity-key",
        help="主体资料：按名称建档，可选补录识别码或任意证件类型与号码"))
    # --db 已由 add_work 挂上（可选：不给就用工作区）
    sp.add_argument("--entity-id", type=int, help="要补录的主体编号")
    sp.add_argument("--uscc", help="统一社会信用代码")
    sp.add_argument("--bank-no", dest="bank_no", help="行内编号")
    sp.add_argument("--id-card", dest="id_card", help="身份证号（自然人主体）")
    sp.add_argument("--doc-type", dest="doc_type", help="其他证件类型，如护照、港澳居民居住证")
    sp.add_argument("--doc-number", dest="doc_number", help="证件号码（可不填；填写时需注明类型）")
    sp.add_argument("--add", action="store_true", help="按名字新建主体；识别码和证件信息可选")
    sp.add_argument("--name", help="主体名（配 --add 用）")
    sp.add_argument("--entity-type", dest="entity_type",
                    choices=["legal_person", "natural_person"], default="legal_person")
    sp.add_argument("--unavailable", action="store_true",
                    help="可选：记下目前没有识别码的原因；不影响继续工作")
    sp.add_argument("--reason", help="拿不到的原因（配 --unavailable 用）")
    sp.add_argument("--batch", help="把事件记到哪个批次（默认自动找锚点）")
    sp.add_argument("--by", default="user")
    sp.add_argument("--out", help="报告输出目录（identifier_report.md）")

    sp = add_work(sub.add_parser(
        "db-merge",
        help="同键勾选覆盖（T-19）：库里已有同键 → 出可勾选表，不勾不动；"
             "勾了写新行、旧行标 superseded_by（历史不丢）"))
    # --db 已由 add_work 挂上（可选：不给就用工作区）
    sp.add_argument("--incoming", help="新来的键值对 JSON（与 db-ingest --profile 同格式）")
    sp.add_argument("--entity-id", type=int, dest="entity_id", help="这批值属于哪个主体")
    sp.add_argument("--name", help="按主体名定位（库里必须已有）")
    sp.add_argument("--uscc", help="按统一社会信用代码定位（库里必须已有 → 不新建）")
    sp.add_argument("--bank-no", dest="bank_no", help="按行内编号定位")
    sp.add_argument("--id-card", dest="id_card", help="按身份证号定位")
    sp.add_argument("--plan", help="第一步产出的 merge_plan.json；带着它执行勾选")
    sp.add_argument("--batch", help="事件记到哪个批次（默认自动找锚点）")
    sp.add_argument("--select", help="勾选：接受新值并覆盖，如 1,3,5")
    sp.add_argument("--select-all", dest="select_all", action="store_true",
                    help="全选（**不含**高风险五类与已交付字段，那两类必须逐个勾）")
    sp.add_argument("--new", action="append",
                    help="改后入库：编号=你输入的值，如 --new 2=广东众森实业发展有限公司")
    sp.add_argument("--discard", action="append", help="丢弃：编号（不写库，但写 fact_discarded 事件）")
    sp.add_argument("--defer", action="append", help="留空待查：编号（不动，进 review_queue）")
    sp.add_argument("--keep", action="append", help="保留库里原值：编号（裁定 took_existing）")
    sp.add_argument("--include-delivered", dest="include_delivered", action="store_true",
                    help="允许勾选已交付过产物的字段（默认拒绝：覆盖不改已出去的文件）")
    sp.add_argument("--add-new", dest="add_new", action="store_true",
                    help="库里还没有的键也一并新增（默认只报告，入库请用 db-ingest）")
    sp.add_argument("--by", default="user")
    sp.add_argument("--out", help="报告输出目录（merge_plan.md / merge_plan.json）")

    sp = add_work(sub.add_parser(
        "db-migrate",
        help="把工作流数据库从 SCHEMA v1 升到 v2（重建 9 张表 + 建 14 张新表），出人看的报告"))
    sp.add_argument("--dry-run", action="store_true",
                    help="预演：只统计将要处理的数据量，不改动数据库")
    sp.add_argument("--day", help="迁移日 YYYYMMDD（给历史批次编号用，默认今天）")
    sp.add_argument("--out", help="报告输出目录（migration_report.md + migration_map.csv）")

    # ------------------------------------------------------------------ ocr
    sp = sub.add_parser("ocr", help="OCR images and scanned PDFs -> text/tables/Word")
    sp.add_argument("input", nargs="+")
    sp.add_argument("--out")
    sp.add_argument("--pages", help="PDF page ranges, e.g. 1-2,5")
    sp.add_argument("--scale", type=float, default=2.5, help="PDF raster scale (higher = finer)")
    sp.add_argument("--preprocess", default="auto", choices=["auto", "none", "gray", "threshold", "enhance"])
    sp.add_argument("--layout", default="auto", choices=["auto", "rows"], help="text layout mode")
    sp.add_argument("--no-tables", dest="tables", action="store_false", default=True,
                    help="skip table-grid reconstruction")
    sp.add_argument("--emit-text", dest="emit_text", action="store_true", default=True)
    sp.add_argument("--no-text", dest="emit_text", action="store_false")
    sp.add_argument("--emit-json", dest="emit_json", action="store_true", default=True)
    sp.add_argument("--no-json", dest="emit_json", action="store_false")
    sp.add_argument("--emit-excel", dest="emit_excel", action="store_true", default=True)
    sp.add_argument("--no-excel", dest="emit_excel", action="store_false")
    sp.add_argument("--emit-markdown", action="store_true", default=False)
    sp.add_argument("--emit-word", action="store_true", default=False, help="also write an editable .docx")
    sp.add_argument("--emit-searchable-pdf", action="store_true", default=False,
                    help="add an invisible text layer (needs reportlab)")

    # ------------------------------------------------------------ file org
    sp = sub.add_parser("dir-list", help="list files with metadata, newest first")
    sp.add_argument("input", nargs="+")
    sp.add_argument("--include", help="comma-separated globs to keep, e.g. '*.xlsx,*.csv'")
    sp.add_argument("--exclude", help="comma-separated globs to skip")
    sp.add_argument("--shallow", action="store_true", help="do not recurse")
    sp.add_argument("--limit", type=int, default=200)
    add_common(sp)

    sp = sub.add_parser("dir-report", help="inventory: categories, sizes, duplicates, stale files, empty dirs")
    sp.add_argument("input", nargs="+")
    sp.add_argument("--out")
    sp.add_argument("--include")
    sp.add_argument("--exclude")
    sp.add_argument("--shallow", action="store_true")
    sp.add_argument("--stale-days", type=float, default=365)
    sp.add_argument("--duplicates", dest="duplicates", action="store_true", default=True)
    sp.add_argument("--no-duplicates", dest="duplicates", action="store_false")
    sp.add_argument("--emit-json", dest="emit_json", action="store_true", default=True)
    sp.add_argument("--no-json", dest="emit_json", action="store_false")
    sp.add_argument("--emit-markdown", dest="emit_markdown", action="store_true", default=True)
    sp.add_argument("--no-markdown", dest="emit_markdown", action="store_false")
    sp.add_argument("--emit-csv", dest="emit_csv", action="store_true", default=True)
    sp.add_argument("--no-csv", dest="emit_csv", action="store_false")

    sp = sub.add_parser("dedupe", help="find (and optionally remove) duplicate files by content hash")
    sp.add_argument("input", nargs="+")
    sp.add_argument("--out")
    sp.add_argument("--include")
    sp.add_argument("--exclude")
    sp.add_argument("--shallow", action="store_true")
    sp.add_argument("--keep", default="oldest",
                    choices=["oldest", "newest", "shortest-path", "longest-path"])
    sp.add_argument("--mode", default="trash", choices=["trash", "delete"],
                    help="trash moves copies into the artifact dir; delete is permanent")
    sp.add_argument("--apply", action="store_true", help="actually act (default is a dry run)")

    sp = sub.add_parser("rename", help="batch rename with a template or regex, dry-run by default")
    sp.add_argument("input", nargs="+")
    sp.add_argument("--out")
    sp.add_argument("--template", help="e.g. '{date}_{nnn}' or '报表_{year}{month}_{name}'")
    sp.add_argument("--prefix", default="")
    sp.add_argument("--suffix", default="")
    sp.add_argument("--find", help="regex applied to the existing stem")
    sp.add_argument("--replace", default="")
    sp.add_argument("--include")
    sp.add_argument("--exclude")
    sp.add_argument("--shallow", action="store_true")
    sp.add_argument("--keep-extension", dest="keep_extension", action="store_true", default=True)
    sp.add_argument("--no-keep-extension", dest="keep_extension", action="store_false")
    sp.add_argument("--apply", action="store_true", help="actually rename (default is a dry run)")

    sp = sub.add_parser("organize", help="sort files into folders by type/date/extension, dry-run by default")
    sp.add_argument("input", help="source directory")
    sp.add_argument("--dest", help="destination directory (default: <input>/整理后_Organized)")
    sp.add_argument("--by", default="type", choices=["type", "date", "year-month", "extension", "first-letter"])
    sp.add_argument("--mode", default="move", choices=["move", "copy"])
    sp.add_argument("--include")
    sp.add_argument("--exclude")
    sp.add_argument("--shallow", action="store_true")
    sp.add_argument("--out")
    sp.add_argument("--apply", action="store_true", help="actually move/copy (default is a dry run)")

    sp = sub.add_parser("archive", help="zip a set of files")
    sp.add_argument("input", nargs="+")
    sp.add_argument("--out")
    sp.add_argument("--name")
    sp.add_argument("--include")
    sp.add_argument("--exclude")
    sp.add_argument("--no-compress", dest="compress", action="store_false", default=True)

    # ---------------------------------------------------------- diagnostics
    sp = sub.add_parser("capabilities", help="report which libraries and engines are usable here")
    sp.add_argument("--json", action="store_true", help="print the raw report")

    return p


def _dispatch(args) -> Result:
    cmd = args.command
    if cmd == "capabilities":
        return _capabilities(args)
    if cmd == "work":
        from .workroot import cmd_work
        return cmd_work(args)
    if cmd == "inspect":
        from .data_ops import cmd_inspect
        return cmd_inspect(args)
    if cmd == "clean":
        from .data_ops import cmd_clean
        return cmd_clean(args)
    if cmd == "profile":
        from .data_ops import cmd_profile
        return cmd_profile(args)
    if cmd == "pivot":
        from .data_ops import cmd_pivot
        return cmd_pivot(args)
    if cmd == "chart":
        from .data_ops import cmd_chart
        return cmd_chart(args)
    if cmd == "report":
        from .data_ops import cmd_report
        return cmd_report(args)
    if cmd == "merge":
        from .data_ops import cmd_merge
        return cmd_merge(args)
    if cmd == "compare":
        from .data_ops import cmd_compare
        return cmd_compare(args)
    if cmd == "data-convert":
        from .data_ops import cmd_convert
        return cmd_convert(args)
    if cmd == "extract":
        from .doc_read import cmd_extract
        return cmd_extract(args)
    if cmd == "build":
        from .doc_build import cmd_build
        return cmd_build(args)
    if cmd == "doc-convert":
        from .doc_build import cmd_convert as cmd_doc_convert
        return cmd_doc_convert(args)
    if cmd == "pdf":
        from .doc_build import cmd_pdf
        return cmd_pdf(args)
    if cmd == "fill":
        from .doc_build import cmd_fill
        return cmd_fill(args)
    if cmd == "fillmap":
        from .doc_fill import cmd_fillmap
        return cmd_fillmap(args)
    if cmd == "to-docx":
        from .doc_build import cmd_to_docx
        return cmd_to_docx(args)
    if cmd == "db-ingest":
        from .harness import cmd_db_ingest
        return cmd_db_ingest(args)
    if cmd == "db-absorb":
        from .absorb import cmd_db_absorb
        return cmd_db_absorb(args)
    if cmd == "db-roles":
        from .harness import cmd_db_roles
        return cmd_db_roles(args)
    if cmd == "db-propose":
        from .harness import cmd_db_propose
        return cmd_db_propose(args)
    if cmd == "db-rules-export":
        from .rule_pack import cmd_db_rules_export
        return cmd_db_rules_export(args)
    if cmd == "db-rules-import":
        from .rule_pack import cmd_db_rules_import
        return cmd_db_rules_import(args)
    if cmd == "db-rule":
        from .harness import cmd_db_rule
        return cmd_db_rule(args)
    if cmd == "db-rule-disable":
        from .harness import cmd_db_rule_disable
        return cmd_db_rule_disable(args)
    if cmd == "db-fill":
        from .harness import cmd_db_fill
        return cmd_db_fill(args)
    if cmd == "db-report":
        from .harness import cmd_db_report
        return cmd_db_report(args)
    if cmd == "db-review":
        from .harness import cmd_db_review
        return cmd_db_review(args)
    if cmd == "db-show":
        from .harness import cmd_db_show
        return cmd_db_show(args)
    if cmd == "db-trace":
        from .harness import cmd_db_trace
        return cmd_db_trace(args)
    if cmd == "db-migrate":
        from .migrate_v2 import cmd_db_migrate
        return cmd_db_migrate(args)
    if cmd == "db-entity-key":
        from .entity_key import cmd_db_entity_key
        return cmd_db_entity_key(args)
    if cmd == "db-merge":
        from .merge_v2 import cmd_db_merge
        return cmd_db_merge(args)
    if cmd == "ocr":
        from .ocr_ops import cmd_ocr
        return cmd_ocr(args)
    if cmd == "dir-list":
        from .file_ops import cmd_list
        return cmd_list(args)
    if cmd == "dir-report":
        from .file_ops import cmd_report
        return cmd_report(args)
    if cmd == "dedupe":
        from .file_ops import cmd_dedupe
        return cmd_dedupe(args)
    if cmd == "rename":
        from .file_ops import cmd_rename
        return cmd_rename(args)
    if cmd == "organize":
        from .file_ops import cmd_organize
        return cmd_organize(args)
    if cmd == "archive":
        from .file_ops import cmd_archive
        return cmd_archive(args)
    raise OfficeKitError(f"unknown command: {cmd}")


def _check_dsh_workroot(args) -> None:
    """DSH 会话首次办公前先认领目录；普通问答与只读查看不受影响。"""
    if not os.environ.get("DSH_SESSION_ID"):
        return
    if args.command in {"capabilities", "inspect", "dir-list", "work"}:
        return

    from .workroot import MARKER_NAME, find_workroot

    explicit_work = getattr(args, "work", None)
    explicit_db = getattr(args, "db", None)
    if explicit_work:
        root = Path(explicit_work).expanduser().resolve()
        found = root if (root / MARKER_NAME).is_file() else None
    elif explicit_db:
        match = find_workroot(Path(explicit_db).expanduser().resolve().parent)
        found = match[0] if match else None
    else:
        match = find_workroot(Path.cwd())
        found = match[0] if match else None

    if found is None:
        raise OfficeKitError(
            "本次办公还没有明确工作区，未执行该命令。先用 office.py work --suggest 看建议路径；"
            "确定后运行 office.py work --init --at <文件夹>，本次创建须经 DSH 批准。"
            "普通聊天、知识问答及 inspect/dir-list 查看不受影响。"
        )

    for option in ("out", "dest"):
        selected = getattr(args, option, None)
        if selected and not Path(selected).expanduser().resolve().is_relative_to(found):
            raise OfficeKitError(
                f"--{option} 指向办公工作区之外：{selected}。"
                f"本次工作区是 {found}；请改用其中的目录，或另选工作区。"
            )


def _capabilities(args) -> Result:
    res = Result("capabilities")
    report: dict[str, object] = {"python": sys.version.split()[0], "executable": sys.executable}

    def check(name: str, module: str, *, extra=None) -> None:
        entry: dict[str, object] = {}
        try:
            mod = __import__(module)
            entry["available"] = True
            entry["version"] = getattr(mod, "__version__", None)
        except Exception as exc:  # noqa: BLE001
            entry["available"] = False
            entry["error"] = f"{type(exc).__name__}: {exc}"
        if extra and entry.get("available"):
            try:
                entry.update(extra())
            except Exception as exc:  # noqa: BLE001
                entry["probe_error"] = f"{type(exc).__name__}: {exc}"
        report[name] = entry

    check("pandas", "pandas")
    check("openpyxl", "openpyxl")
    check("xlsxwriter", "xlsxwriter")
    check("pyarrow", "pyarrow")
    check("python-docx", "docx")
    check("docxtpl", "docxtpl")
    check("python-pptx", "pptx")
    check("pypdf", "pypdf")
    check("pypdfium2", "pypdfium2")
    check("pdfplumber", "pdfplumber")
    check("rapidocr", "rapidocr_onnxruntime")
    check("opencv", "cv2")
    check("matplotlib", "matplotlib")
    check("reportlab", "reportlab")
    check("xlrd", "xlrd")

    def cjk_fonts():
        from .common import CJK_FONT_CANDIDATES, configure_matplotlib

        _, chosen = configure_matplotlib()
        return {"selected": chosen, "candidates_present": list(CJK_FONT_CANDIDATES)}

    check("matplotlib-cjk", "matplotlib", extra=cjk_fonts)

    def com():
        from . import office_com

        return {"apps": office_com.availability_report()}

    check("ms-office-com", "win32com.client", extra=com)

    def browser():
        from .doc_build import find_headless_browser

        found = find_headless_browser()
        return {"found": found} if found else {"found": None, "note": "HTML->PDF falls back to matplotlib"}

    report["headless-browser"] = browser()

    def ocr_ready():
        from . import ocr_ops

        return ocr_ops.engine_info()

    try:
        report["ocr-engine"] = ocr_ready()
    except Exception as exc:  # noqa: BLE001
        report["ocr-engine"] = {"available": False, "error": str(exc)}

    def fingerprint():
        from .fingerprint import tool_fingerprint

        return tool_fingerprint()

    report["tool-fingerprint"] = fingerprint()

    res.data["report"] = report
    res.data["summary"] = {
        "data_analysis": all(report.get(k, {}).get("available") for k in ("pandas", "openpyxl")),
        "ocr_ready": bool(report.get("ocr-engine", {}).get("installed")),
        "office_com_ready": any(
            v.get("available") for v in (report.get("ms-office-com", {}).get("apps") or {}).values()
        ),
        "high_fidelity_pdf": bool(report.get("headless-browser", {}).get("found")),
    }
    return res


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    real_stdout = sys.stdout
    # Native libraries (ONNX Runtime, pywin32, browsers) may print to stdout at
    # import or inference time.  Funnel everything through stderr while the work
    # runs, so the JSON envelope is the only thing the agent ever parses.
    sys.stdout = sys.stderr
    try:
        try:
            _check_dsh_workroot(args)
            result = _dispatch(args)
            payload, code = result.to_json(ok=True), 0
        except OfficeKitError as exc:
            res = Result(getattr(args, "command", "unknown"))
            payload, code = res.to_json(ok=False, error=str(exc)), 2
        except KeyboardInterrupt:
            res = Result(getattr(args, "command", "unknown"))
            payload, code = res.to_json(ok=False, error="interrupted"), 130
        except Exception as exc:  # noqa: BLE001
            res = Result(getattr(args, "command", "unknown"))
            res.data["exception"] = type(exc).__name__
            if "--debug" in (argv or sys.argv):
                res.data["traceback"] = traceback.format_exc()
            payload, code = res.to_json(ok=False, error=f"{type(exc).__name__}: {exc}"), 1
    finally:
        sys.stdout = real_stdout
    print(payload)
    return code


if __name__ == "__main__":
    raise SystemExit(main())
