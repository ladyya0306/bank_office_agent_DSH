# toolV2 维护入口

适用：用户明确要求修复或扩展程序。普通填表仍执行`skills/office-suite-v2/SKILL.md`，不会因一份文件失败而自行改共享程序。

需要理解运行目录与开发工作树、安装包、启用及回退时，先看[完整使用与部署说明](../release-docs/交给DSH的升级提示词.md)。选择会话目录不会改变当前DSH的填表引擎路径。

## 先理解这四件事

1. toolV2是已有Word/Excel引擎之上的固定填表流程，不是15份测试题的答案库。
2. DSH模型负责不明确位置的语义映射；程序负责来源缓存、主体事实、原生确认、文件生成、校验和恢复。
3. 借款人/保证人是当前主要业务语境，不能自动推广成产品主数据或300主体批量任务。扩展应复用引擎，增加明确任务模式或来源读取能力。
4. 当前事实在`../release-docs/maintenance-state.md`。读完本页与状态即可开始定位；不要先读全部历史报告或整仓库源码。

## 到哪里修改

| 问题或需求 | 优先阅读 |
|---|---|
| 临时脚本识别结果、用户确认、保存复用 | `workflow/learning.py`、`workflow/source_evidence.py`；[说明](../release-docs/已确认脚本与填写方法.md) |
| 来源识别、主体归属 | `workflow/source.py`、`office_kit/absorb.py`、`office_kit/fact_catalog.py` |
| 来源冲突 | `workflow/source_conflicts.py` |
| 模板空位、字段对应、文件名角色 | `office_kit/template_slots.py`、`workflow/mapping.py`、`office_kit/target_validation.py` |
| 重复提问、复用回答 | `workflow/review.py`、`office_kit/fill_decisions.py` |
| 金额单位、实际Word/Excel输出 | `office_kit/value_fit.py`、`office_kit/harness.py`、`office_kit/doc_fill.py` |
| 状态、缓存、局部重做 | `workflow/runner.py`、`workflow/storage.py`、`workflow/timing.py` |
| 模型读到的分页、来源与产物 | `workflow/mapping_view.py` |
| 原生弹窗、取消与交付 | `dsh-plugin/controller.mjs`、`dsh-plugin/index.mjs`、`dsh-plugin/render.mjs`、`workflow/delivery.py` |
| 接入或启动日志 | `install.py`、`../deepseek-harness-local/start.js`、`../deepseek-harness-local/runtime-logging.cjs` |
| PDF/OCR扩展的已有底层 | `office_kit/cli.py`、`office_kit/ocr_ops.py`；先查入口实际调用，不只看参数名 |

## 一项改动怎样完成

- **定范围**：把用户需求转成输入、输出、验收条件和不变项。能从材料和代码确定的事自行处理，只询问影响业务含义的缺失信息。
- **查现状**：记录`git status --short`、分支、HEAD和现有相关测试。已有用户改动保留；在干净代码工作树的功能分支施工，不将混有客户资料的`D:\DSH`当作开发仓库整体提交。
- **建立证据**：缺陷先做最小合成复现；新功能先定义可检查输出。例如300家公司应得到300份结果，逐份公司标识匹配，甲公司的利率不出现在乙公司的文件中。
- **小步修改**：先定位根因，再复用现有函数和工具。不要重写一套填表器，不以任意默认值、放宽权限、清库或大量留空换取完成状态。
- **验证**：先跑受影响的测试，修复后检查真实生成的合成DOCX/XLSX。共享解析、状态、确认或写入器改变时补跑全量；纯文档变更只查内容与链接。不在代码未变化、相关测试已通过时反复跑同一套测试。
- **保存**：相关检查通过后提交明确文件并推送功能分支；不能把提交或推送成功写成运行版已经启用。验收文档写实际命令、通过/跳过/失败、产物、耗时和限制。

程序升级涉及解析版本、缓存或执行签名时，先查`runner.py`、`source.py`和`review.execution_signature`的现有机制。变化确实影响旧缓存时才升相应版本，并测“只重做受影响文件”；不要靠清空数据库或每次全部重跑生效。

## 可复制的验证命令

以下在仓库根目录的PowerShell执行，使用已准备好的开发依赖；缺依赖要如实报告，不能把导入失败称为测试通过。

```powershell
$env:PYTHONUTF8='1'
$env:PYTHONPATH=(Resolve-Path toolV2).Path
python -m pytest -q toolV2/tests
node --test toolV2/tests/plugin-controller.mjs toolV2/tests/plugin-delivery-render.mjs toolV2/tests/plugin-integration.mjs toolV2/tests/plugin-process-tree.mjs toolV2/tests/plugin-workspaces.mjs
```

只改启动器时，可先执行：

```powershell
node --test deepseek-harness-local/tests/runtime-logging.cjs
```

改变审批插件时另执行它已有的`selftest.mjs`。各次测试数量以实际输出为准，不照抄历史数字。

注意：日常DSH的`office_fill_task`可能仍指向`D:\DSH\toolV2`，不会因切换源码工作树自动使用候选代码。测试候选代码要核对实际模块路径；DSH集成复验还要核对接入的`tool_root`。不能在旧程序上跑通后宣称新程序已验证。

模型开发测试仅使用合成资料。未知模板测试应先冻结版本、隔离参考结果，结束后再评分；不把参考答案喂给执行流程。记录首次生成与旧任务复用两个口径，token须区分总量与可取得的缓存明细。

## Git与GitHub

- 本仓库公开，只提交程序、公开文档和合成测试。提交前检查`git diff`、`git diff --cached`和文件清单；排除客户资料、数据库、`.env`、会话、日志、凭据、原始测试报告。`.gitignore`不是检查内容的替代品。
- 在当前维护版本上创建功能分支。取得远端更新后确认历史关系，正常推送当前功能分支；不强推、不改已有验收标签、不自动覆盖`main`。发现远端新增提交先比较并处理，不能以强推跳过。
- 一次成功推送后比较`git rev-parse HEAD`与`git ls-remote --heads origin <当前分支>`的提交值。鉴权失败时保留本地提交并说明原因，不把密钥贴入命令或文档，不谎报上传成功。
- 可执行的保存顺序为：检查改动→明确路径`git add`→检查暂存内容→`git commit`→`git push -u origin HEAD`→核对远端。推送须在本次用户授权的范围内。

## 启用与下一次交接

若尚未授权启用，先交付可复核代码和测试结果，再确认启用；已有授权则继续完成。启用前备份受影响程序和必要接入配置，核对候选与运行文件，检查正在执行的任务。程序回退不等于数据库回退；不复制开发数据库覆盖业务数据库，不覆盖整个运行目录。

本机配置与接入沿用`install.py`和已有部署方法，不能为了维护方便放开目录、网络或批准边界。普通业务确认卡不能被当作修改共享程序的授权。

完成或中断前更新`../release-docs/maintenance-state.md`：目标、当前提交/分支、实际改动、测试、已启用或待启用、未解决问题和下一个具体动作。新会话读该文件继续，不要求用户重讲全部历史。
