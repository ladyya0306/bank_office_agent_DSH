/**
 * show_approvals.mjs —— 从**会话记录**里把「批准」这件事翻出来给人看。
 *
 *   node scripts/show_approvals.mjs              # 最近一次会话
 *   node scripts/show_approvals.mjs --types      # 先看有哪些事件类型
 *   node scripts/show_approvals.mjs --session <会话目录或文件>
 *
 * 为什么要它：会话记录是 `.jsonl.zstd`（压缩过的），用记事本打不开。
 * 批准这件事的全部证据都在里面：`approval/asked`（谁问了什么）、
 * `approval/decided`（人点了什么）、以及工具调用的最终结果（到底执行没有）。
 *
 * ⚠️ 边界：它只**读**记录、不改任何东西；记录本身没做过防篡改校验，
 *    所以只能叫"可供追查"，不能说"不可篡改"。
 */
import { existsSync, readFileSync, readdirSync, statSync } from 'node:fs';
import { zstdDecompressSync } from 'node:zlib';
import { dirname, join, resolve } from 'node:path';
import { homedir } from 'node:os';
import { fileURLToPath } from 'node:url';

const localHome = resolve(dirname(fileURLToPath(import.meta.url)), '..',
  'deepseek-harness-local', 'workspace');
const DSH_HOME = process.env.DSH_HOME ||
  (existsSync(join(localHome, 'sessions')) ? localHome : join(homedir(), '.dsh'));

const ZSTD_MAGIC = Buffer.from([0x28, 0xb5, 0x2f, 0xfd]);

/**
 * 把一整份会话记录解出来。
 *
 * ⚠️ 这个文件是**一串首尾相接的 zstd 帧**（每落一段追一帧；实测一个 635KB 的文件里有
 * **316 帧**）。Node 的流式解压只肯解第一帧，`zstdDecompressSync` 也一样——所以必须
 * 自己按魔数切帧、逐帧解。切点不一定都在真帧尾（压缩数据里也可能出现魔数），
 * 所以每一刀都用「解不开就往后找下一刀」的办法试。
 */
function decodeFrames(buf) {
  const starts = [];
  for (let i = 0; i + 4 <= buf.length; i++) {
    if (buf.compare(ZSTD_MAGIC, 0, 4, i, i + 4) === 0) starts.push(i);
  }
  if (starts.length === 0 || starts[0] !== 0) return null;   // 不是这种格式
  starts.push(buf.length);
  const parts = [];
  let k = 0;
  while (k < starts.length - 1) {
    let advanced = false;
    for (let j = k + 1; j < starts.length; j++) {
      try {
        parts.push(zstdDecompressSync(buf.subarray(starts[k], starts[j])).toString('utf8'));
        k = j;
        advanced = true;
        break;
      } catch { /* 没切在帧尾，继续往后找 */ }
    }
    if (!advanced) break;                                    // 尾部残缺帧：丢掉，不假装读全了
  }
  return parts.join('');
}

/** 找最近的会话记录文件。 */
function newestSessionFile() {
  const base = join(DSH_HOME, 'sessions');
  const found = [];
  const walk = (dir, depth) => {
    let entries;
    try { entries = readdirSync(dir, { withFileTypes: true }); } catch { return; }
    for (const e of entries) {
      const p = join(dir, e.name);
      if (e.isDirectory() && depth < 2) walk(p, depth + 1);
      else if (e.isFile() && e.name.startsWith('session.') && e.name.endsWith('.zstd')) {
        try { found.push({ p, mtime: statSync(p).mtimeMs }); } catch { /* 跳过读不到的 */ }
      }
    }
  };
  walk(base, 0);
  found.sort((a, b) => b.mtime - a.mtime);
  return found[0]?.p ?? null;
}

const argv = process.argv.slice(2);
const wantTypes = argv.includes('--types');
const wantList = argv.includes('--list');
const idx = argv.indexOf('--session');
let file = idx >= 0 ? argv[idx + 1] : null;
if (file && !file.endsWith('.zstd')) file = join(file, 'session.v3.jsonl.zstd');
if (!file) file = newestSessionFile();
if (!file) { console.error('找不到会话记录文件'); process.exit(2); }

/** 读一个会话文件 → 事件数组（读不动返回 null）。 */
function eventsOf(path) {
  try {
    const text = decodeFrames(readFileSync(path));
    if (text === null) return null;
    const out = [];
    for (const line of text.split('\n')) {
      const s = line.trim();
      if (s === '') continue;
      try { out.push(JSON.parse(s)); } catch { /* 尾部半行忽略 */ }
    }
    return out;
  } catch { return null; }
}

if (wantList) {
  const base = join(DSH_HOME, 'sessions');
  const all = [];
  const walk = (dir, depth) => {
    let entries;
    try { entries = readdirSync(dir, { withFileTypes: true }); } catch { return; }
    for (const e of entries) {
      const p = join(dir, e.name);
      if (e.isDirectory() && depth < 2) walk(p, depth + 1);
      else if (e.isFile() && e.name.startsWith('session.') && e.name.endsWith('.zstd')) all.push(p);
    }
  };
  walk(base, 0);
  const rows = [];
  for (const p of all) {
    const st = statSync(p);
    const evs = eventsOf(p);
    rows.push({
      p, mtime: st.mtimeMs, size: st.size, evs: evs?.length ?? -1,
      approvals: (evs ?? []).filter((e) => String(e.type ?? '').startsWith('approval/')).length,
    });
  }
  rows.sort((a, b) => b.mtime - a.mtime);
  console.log(`DSH_HOME = ${DSH_HOME}`);
  console.log('最近改动的会话记录（前 12 个）：');
  for (const r of rows.slice(0, 12)) {
    console.log(`  ${new Date(r.mtime).toLocaleString()}  ${String(r.size).padStart(9)}B  事件 ${String(r.evs).padStart(5)}  批准 ${r.approvals}  ${r.p}`);
  }
  process.exit(0);
}

const events = eventsOf(file);
if (events === null) { console.error(`读不动：${file}`); process.exit(2); }

console.log(`会话记录：${file}`);
console.log(`事件数：${events.length}`);

if (wantTypes) {
  const counts = new Map();
  for (const ev of events) counts.set(ev.type ?? '(无类型)', (counts.get(ev.type ?? '(无类型)') ?? 0) + 1);
  console.log('\n事件类型（次数）：');
  for (const [type, n] of [...counts].sort()) console.log(`  ${n.toString().padStart(5)}  ${type}`);
  process.exit(0);
}

const brief = (v, limit = 220) => {
  const s = typeof v === 'string' ? v : JSON.stringify(v);
  return s === undefined ? '' : (s.length <= limit ? s : s.slice(0, limit) + ' …');
};

console.log('\n=== 批准相关 ===');
let approvals = 0;
const askedById = new Map();
const alwaysCallIds = new Set();
for (const ev of events) {
  const t = String(ev.type ?? '');
  if (!t.startsWith('approval/')) continue;
  approvals++;
  const d = ev.data ?? {};
  if (t === 'approval/asked') {
    askedById.set(d.id, d.callId);
    console.log(`  #${ev.seq} 询问  工具=${d.toolName ?? d.tool ?? '?'}  请求号=${d.id ?? '?'}  轮次=${d.turn ?? '?'}`);
  } else if (t === 'approval/policy' && d.bankGrant) {
    alwaysCallIds.add(d.bankGrant.callId);
    console.log(`  #${ev.seq} 本会话同类操作一直允许  类别=${d.bankGrant.scope}  工作区=${d.bankGrant.cwd}`);
  } else if (t === 'approval/decided') {
    const always = alwaysCallIds.has(askedById.get(d.id));
    console.log(`  #${ev.seq} 决定  结果=${always ? '本次放行，且本会话同类持续允许' : d.outcome ?? '?'}  请求号=${d.id ?? '?'}`);
  } else {
    console.log(`  #${ev.seq} ${t}  ${brief(d, 160)}`);
  }
}
if (approvals === 0) console.log('  （没有批准事件：本会话没有触发过批准）');

console.log('\n=== 工具调用与结果 ===');
let tools = 0;
for (const ev of events) {
  const t = String(ev.type ?? '');
  if (!/^tool/i.test(t) && !/tool\//.test(t)) continue;
  tools++;
  const d = ev.data ?? {};
  const name = d.name ?? d.toolName ?? d.tool ?? '?';
  const kind = /result/i.test(t) ? '结果' : '调用';
  const payload = /result/i.test(t) ? (d.result ?? d.content ?? d) : (d.arguments ?? d.args ?? d);
  console.log(`  #${ev.seq} ${kind}  ${name}  ${brief(payload, 200)}`);
}
if (tools === 0) console.log('  （没有工具事件）');

