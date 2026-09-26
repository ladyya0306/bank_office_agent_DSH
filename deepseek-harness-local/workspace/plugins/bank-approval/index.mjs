/**
 * bank-approval —— 把「这一次要不要动手」交给 DSH 自带的批准框。
 *
 * 它在 `tools/pre-execute` 分类，并用 DSH 原生批准框接收一次性决定。
 * “本会话同类操作一直允许”记在 DSH 会话记录中；新会话不继承。
 *
 * 三件事分得很清：
 *   · **范围**（能写哪里）＝ 沙箱模式（`workspace-write` 预设）管的，不是本插件管的；
 *   · **这一次要不要动手** ＝ 本插件管的（弹框，人点）；
 *   · **改完能不能交付** ＝ 办公工具自己的四道门禁管的。
 * 工作区标记只说明范围，**不代表使用者批准过某一次操作**，所以本插件不读它。
 *
 * 判断口径（第一版，规则刻意短，拿不准就走一次批准）：
 *   · 只读/展示类工具            → 继续（不弹框）
 *   · 写文件的工具               → 批准（卡片写明：全路径 + 新建还是覆盖）
 *   · 命令工具里的**核实过的只读命令** → 继续
 *   · 命令工具的其他调用（含跑新脚本） → 批准（卡片写明：完整命令 + 工作目录）
 *   · 办公工具：只读子命令 → 继续；会改数据/出交付物的子命令 → 批准
 *   · 没见过的工具               → 批准（不许悄悄放行）
 *   · 少数不可逆且明确禁止的动作 → 拒绝（写明命中哪条、可用什么替代）
 */
import { existsSync } from 'node:fs';

export const name = 'bank-approval';
export const inject = ['approval'];

/** 只读或无副作用的工具：继续，不打扰。表里没有的工具一律走批准。 */
const FREE_TOOLS = new Set([
  'read', 'read_image', 'glob', 'grep', 'web_search', 'web_fetch', 'skill',
  'office_fill_review',
  'todo_write', 'present', 'ask_user_question', 'list_agents', 'job_list',
  'job_output', 'job_kill', 'get_goal', 'create_goal', 'update_goal',
  'send_message', 'interrupt_agent', 'subagent', 'subagent_fork', 'workflow',
  'ralph', 'exit_plan_mode',
]);

/** Known workspace-scoped office operation; its own resolver/input checks remain authoritative. */
const OFFICE_TASK_TOOLS = new Set(['office_fill_task']);

/** 写文件的工具：一律批准。 */
const FILE_TOOLS = new Set(['write', 'edit', 'str_replace', 'create_file', 'str_replace_editor']);

/** 跑命令的工具：按命令内容判。 */
const SHELL_TOOLS = new Set(['pwsh', 'bash', 'shell', 'cmd']);

/** 命令第一段允许出现的只读命令（**白名单**，不是黑名单）。 */
const READONLY_HEAD = new Set([
  'get-childitem', 'gci', 'ls', 'dir', 'get-content', 'gc', 'cat', 'type',
  'test-path', 'select-string', 'sls', 'findstr', 'measure-object', 'get-item',
  'gi', 'get-location', 'pwd', 'resolve-path', 'split-path', 'join-path',
  'get-command', 'get-help', 'get-member', 'get-date', 'get-process', 'ps',
  'get-service', 'get-volume', 'get-psdrive', 'where.exe', 'where', 'which',
  'hostname', 'whoami', 'tree',
]);

/** 管道/分段的后续段允许出现的只读命令。 */
const PIPE_HEAD = new Set([
  'select-object', 'sort-object', 'group-object', 'measure-object', 'where-object',
  'format-table', 'ft', 'format-list', 'fl', 'format-wide', 'out-string',
  'select-string', 'sls', 'get-unique', 'compare-object', 'tee-object', 'convertto-json',
]);

/** 出现这些就算"在跑程序"：不能靠命令名猜只读，一律走批准。 */
const RUNS_PROGRAM = /(^|[\s"'=|&;(\\/])(python[0-9.]*|py|node|npm|npx|pnpm|yarn|git|pip|conda|dotnet|java|javac|go|cargo|ruby|perl|php|powershell|pwsh|cmd|wsl|bash|sh|reg|netsh|schtasks|sc|icacls|takeown)(\.exe)?(\s|$)/i;

/**
 * 办公工具子命令的**参数级分类**（第 29 批，主人已确认口径）。
 *
 * 三层：
 *  · `OFFICE_READ`       —— 只看/预览：免打扰；
 *  · `OFFICE_DEPENDS`    —— **同一个命令两种身份**：带"动手旗标"才请批准，不带就是看；
 *  · `OFFICE_WRITE`      —— 改数据 / 出交付物：一律请批准。
 *  没登记的子命令**一律请批准**（宁可多问一次，不许悄悄放行）。
 *
 * ⚠️ 已确认的一个判断：**`--out` 不改判**——预览类命令顺手写一份报告文件仍免打扰。
 *    理由：它写的是报告，不是改数据、也不是交付物；否则每看一眼都要点卡片，人会被逼到乱点。
 */
const OFFICE_READ = new Set([
  'capabilities', 'inspect', 'profile', 'compare', 'extract',
  'db-show', 'db-trace', 'db-report', 'dir-list', 'db-absorb',
]);

/** 带这些旗标才"动手"；一个都不带就是预览。 */
const OFFICE_DEPENDS = {
  work: ['--init', '--triage'],
  'db-review': ['--answer', '--new-value'],
  'db-merge': ['--select', '--select-all', '--keep', '--discard', '--defer', '--add-new', '--incoming', '--new'],
  'db-entity-key': ['--add', '--uscc', '--bank-no', '--id-card', '--doc-type', '--doc-number', '--unavailable', '--name', '--entity-type'],
  dedupe: ['--apply'],
  rename: ['--apply'],
  organize: ['--apply'],
  // 注意：`db-migrate` 不在这里——它**默认就是升级（动手）**，只有 `--dry-run` 才是预览，
  // 所以它归 OFFICE_WRITE，并在分支里对 `--dry-run` 单判。
};

/** 改数据 / 出交付物 / 签核：一律请批准。 */
const OFFICE_WRITE = new Set([
  'clean', 'pivot', 'chart', 'report', 'merge', 'data-convert', 'build', 'doc-convert',
  'pdf', 'fill', 'fillmap', 'to-docx', 'db-ingest', 'db-roles', 'db-propose', 'db-rule',
  'db-fill', 'db-rules-export', 'db-rules-import', 'db-rule-disable', 'ocr', 'archive', 'db-migrate',
]);

/** 少数不可逆、且本项目已明确禁止的动作：拒绝，并给出替代做法。 */
const FORBIDDEN = [
  { re: /\bpython[0-9.]*\s+-c\b/i, why: '`python -c` 这种"现敲一段代码就跑"的写法',
    fix: '改成先把脚本写成文件（写文件会走批准），再 `python 脚本.py`' },
  { re: /\b(invoke-expression|iex)\b/i, why: '`Invoke-Expression` 把字符串当代码跑',
    fix: '改成写成脚本文件再执行' },
  { re: /\b(curl|wget|iwr|invoke-webrequest|bitsadmin|certutil)\b/i, why: '从网上下载东西',
    fix: '行内环境不外联；需要文件就由使用者放到工作区里' },
  { re: /\b(remove-item)\b[^|;]*\s-(recurse|force)\b/i, why: '递归/强制删除',
    fix: '先列出要删的东西给人看，再逐个删；删产物请走办公工具的清理流程' },
  { re: /\b(rmdir)\s+\/s\b|\b(del)\s+\/[fsq]\b/i, why: '递归删除目录/强制删除',
    fix: '同上' },
  { re: /\b(reg\s+(add|delete)|netsh\s|sc\s+config|schtasks\s+\/create|set-executionpolicy|new-localuser|add-localgroupmember)\b/i,
    why: '改系统设置（注册表/网络/服务/计划任务/执行策略/本机账号）',
    fix: '这类改动不由会话代做，请走行内运维流程' },
  // ⚠️ 第 29 批修误杀：原来写 `\b(format|diskpart)\b`，把**只读的 `Format-Table`/`Format-List`**
  //    也拒了（真实会话里已发生）。磁盘格式化一定是 `format <盘符>:` 这种形状，按形状判。
  { re: /\bformat\s+[a-z]:|\bdiskpart\b|\bformat-volume\b/i,
    why: '格式化/分区操作', fix: '绝不允许在会话里做' },
];

function grantScope(tool, command, sub) {
  if (FILE_TOOLS.has(tool)) return '写文件';
  if (sub) return `办公命令：${sub}`;
  if (SHELL_TOOLS.has(tool)) {
    const match = command.match(RUNS_PROGRAM);
    return match ? `运行程序：${match[2].toLowerCase()}` : '其他命令';
  }
  return `工具：${tool}`;
}

function hasGrant(session, scope, cwd) {
  if (!session || !scope || !cwd) return false;
  for (let seq = session.seq - 1; seq >= 0; seq -= 1) {
    const event = session.eventAt(seq);
    const grant = event?.type === 'approval/policy' ? event.data?.bankGrant : null;
    if (grant?.scope === scope && grant?.cwd === cwd) return true;
  }
  return false;
}

const pendingScopes = new Map();

/**
 * 把命令里**所有** `office.py <子命令>` 找出来，返回**最严**的那个。
 *
 * ⚠️ 第 29 批修缺口：原来只取第一个匹配 —— `office.py capabilities; office.py db-migrate`
 *    会被当成"预览"整体放行。现在**只要有一个是动手层，整体就按动手层判**。
 */
function officeSubcommand(command) {
  const re = /office\.py"?\s+([a-z][a-z0-9-]*)/gi;
  const low = command.toLowerCase();
  let strictest = null;
  let m;
  while ((m = re.exec(command)) !== null) {
    const sub = m[1].toLowerCase();
    if (strictest === null || officeStrictness(sub, low) > officeStrictness(strictest, low)) {
      strictest = sub;
    }
  }
  return strictest;
}

/** 0＝预览（免打扰）／1＝动手（必问）。**同一个命令两种身份**按"动手旗标"判。 */
function officeStrictness(sub, low) {
  if (OFFICE_READ.has(sub)) return 0;
  if (sub === 'db-migrate') return low.includes('--dry-run') ? 0 : 1;   // 默认就是升级
  const need = OFFICE_DEPENDS[sub];
  if (need !== undefined && need.length > 0) return need.some((f) => low.includes(f)) ? 1 : 0;
  return 1;                                   // 动手层 或 未登记 → 一律问
}

/** 命令是不是"核实过的只读"：每段头都在白名单里，且不含重定向/写动作。 */
function isVerifiedReadonly(command) {
  if (/[>]/.test(command)) return false;                    // 任何重定向都算写
  if (RUNS_PROGRAM.test(command)) return false;             // 跑程序 → 不猜，走批准
  const segments = command.split(/\||;|&&|\n/);
  for (const raw of segments) {
    const seg = raw.trim().replace(/^&\s*/, '').trim();
    if (seg === '') continue;
    const head = seg.split(/\s+/)[0].replace(/^["']|["']$/g, '').split(/[\\/]/).pop()
      .replace(/\.exe$/i, '').toLowerCase();
    if (segments.length === 1) {
      if (!READONLY_HEAD.has(head) && !PIPE_HEAD.has(head)) return false;
    } else if (!PIPELINE_SEGMENT_OK(head)) {
      return false;
    }
  }
  return true;
}

/** 多段命令里，第一段用只读表，后续段只允许管道用的那几个。 */
function PIPELINE_SEGMENT_OK(head) {
  return READONLY_HEAD.has(head) || PIPE_ALLOWED.has(head);
}
const PIPE_ALLOWED = new Set([...PIPE_HEAD, 'out-string', 'write-output', 'echo']);

function short(text, limit) {
  const value = String(text ?? '');
  return value.length <= limit ? value : value.slice(0, limit) + ` …（共 ${value.length} 字，已截断）`;
}

/** 给使用者看的"这一步要干什么"。返回 {kind:'next'|'deny'|'ask', reason?}。 */
function classify(toolName, args, cwd) {
  const tool = String(toolName || '').toLowerCase();
  const a = args && typeof args === 'object' ? args : {};

  // The dedicated office tool already validates its work root, inputs, native
  // data state, and output locations. Do not add a second, per-session approval.
  if (OFFICE_TASK_TOOLS.has(tool)) return { kind: 'next' };

  if (FREE_TOOLS.has(tool)) return { kind: 'next' };

  if (FILE_TOOLS.has(tool)) {
    const path = String(a.file_path ?? a.path ?? '');
    if (path === '') {
      return { kind: 'ask', reason: '写文件，但**没给路径**——请先说明要写哪个文件。' };
    }
    let what = '新建';
    try {
      what = existsSync(path) ? '**覆盖修改已有文件**' : '新建文件';
    } catch { /* 读不到就当新建，卡片里写清是"不确定" */ what = '新建（未能确认是否已存在）'; }
    return {
      kind: 'ask',
      scope: grantScope(tool),
      reason: `要写文件（${what}）：\n\n    ${path}\n\n同类范围：写文件。可允许一次，或在本会话当前工作区内一直允许写文件。`,
    };
  }

  if (SHELL_TOOLS.has(tool)) {
    const command = String(a.command ?? a.cmd ?? '');
    if (command === '') return { kind: 'ask', reason: '要执行一条命令，但命令内容是空的。' };

    for (const rule of FORBIDDEN) {
      if (rule.re.test(command)) {
        return {
          kind: 'deny',
          reason: `这一步被**明确禁止**：${rule.why}。\n\n→ 替代做法：${rule.fix}`,
        };
      }
    }

    const sub = officeSubcommand(command);
    if (sub !== null) {
      const low = command.toLowerCase();
      if (officeStrictness(sub, low) === 0) return { kind: 'next' };   // 预览层（最严的那个也是预览）
      return {
        kind: 'ask',
        scope: grantScope(tool, command, sub),
        reason: `要跑办公工具里**会动手**的子命令 \`${sub}\`（一条命令里若串了多个子命令，按最严的那个判）：\n\n`
          + `    ${short(command, 900)}\n\n工作目录：${cwd}\n同类范围：办公命令 ${sub}。可允许一次，或本会话同类操作一直允许。`
          + `若这一步会改数据或出交付物，请先让使用者看过预演/清单再点。`,
      };
    }

    if (isVerifiedReadonly(command)) return { kind: 'next' };

    return {
      kind: 'ask',
      scope: grantScope(tool, command),
      reason: `要执行一条**不在只读白名单里**的命令：\n\n    ${short(command, 900)}\n\n`
        + `工作目录：${cwd}\n同类范围：${grantScope(tool, command)}。可允许一次，或本会话同类操作一直允许。`,
    };
  }

  return {
    kind: 'ask',
    scope: grantScope(tool),
    reason: `要用一个**没有登记过的工具** \`${toolName}\`：\n\n    ${short(JSON.stringify(a), 600)}\n\n`
      + `同类范围：${grantScope(tool)}。可允许一次，或本会话同类操作一直允许。`,
  };
}

/**
 * 挂到工具执行入口。
 *
 * 三条护栏，按重要性排：
 * 1. **会话审批策略是 `never` 时，不请求批准，也不放行需要批准的动作**。
 *    只读动作照常执行；需要批准的动作明确拒绝，并提示使用者切换权限预设。
 * 2. **任何内部错误都不许变成"静默放行"**：出错时改为请求批准，把决定交回使用者。
 * 3. 其余照 `classify` 的判断走。
 */
export function apply(ctx) {
  ctx.on('approval/request', async (request, next) => {
    try {
      const outcome = await next();
      if (outcome !== 'allowed-always') return outcome;
      const session = request?.agent?.session;
      const candidate = pendingScopes.get(request?.callId);
      if (!candidate || candidate.session !== session || !candidate.scope || !candidate.cwd) return 'rejected';
      session.append('approval/policy', {
        policy: 'ask',
        bankGrant: { scope: candidate.scope, cwd: candidate.cwd,
                     callId: request.callId },
      });
      return 'allowed-once';
    } catch (error) {
      ctx.logger?.warn?.(`bank-approval: 无法记录一直允许的决定：${String(error)}`);
      return 'rejected';
    } finally {
      if (request?.callId !== undefined) pendingScopes.delete(request.callId);
    }
  }, true);
  ctx.on('tools/pre-execute', (exec, next) => {
    let verdict;
    try {
      verdict = classify(exec?.name, exec?.arguments, exec?.agent?.session?.header?.cwd ?? '');
    } catch (error) {
      const detail = error instanceof Error ? `${error.name}: ${error.message}` : String(error);
      try { ctx.logger?.warn?.(`bank-approval: 判断出错，改为请求批准：${detail}`); } catch { /* 日志失败不影响判断 */ }
      return { kind: 'ask', reason: `判断这一步时插件内部出错了（${detail}）——**不敢替你放行**，请人工看一眼再决定。` };
    }
    if (verdict.kind === 'next') return next();
    if (verdict.kind === 'deny') return { kind: 'deny', reason: verdict.reason };
    // `never` 无法弹出原生批准框。不能因此把写操作降级放行。
    try {
      if (ctx.get('approval')?.effectivePolicy?.(exec?.agent?.session) === 'never') {
        return {
          kind: 'deny',
          reason: '本会话关闭了批准请求，这一步尚未得到用户同意。请在 DSH 中执行 /permission workspace-write，确认当前会话可以弹批准框后再重试。',
        };
      }
    } catch { /* 读不到策略则交给原生批准服务处理；它缺席时会拒绝。 */ }
    const session = exec?.agent?.session;
    const cwd = session?.header?.cwd ?? '';
    if (hasGrant(session, verdict.scope, cwd)) return next();
    if (exec?.callId !== undefined && verdict.scope) {
      pendingScopes.set(exec.callId, { session, cwd, scope: verdict.scope });
    }
    return { kind: 'ask', reason: verdict.reason };
  });
}
