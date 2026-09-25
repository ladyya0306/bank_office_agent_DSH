/**
 * 新会话「会话目录 → 办公工作区确认」合同测试。
 *
 * 用法：
 *   node tests/session-office-gate.contract.mjs
 *
 * 适配器必须导出 createSessionOfficeGateTestAdapter({ workspaceHome })，返回：
 *   - selectSessionDirectory(sessionId, absolutePath)
 *   - selectOfficeWorkspace(sessionId, absolutePath)
 *   - runUserPromptHook(sessionId, prompt)
 *       （必须实际启动生产 UserPromptSubmit hook 子进程，返回
 *        { exitCode, stdout, stderr }，而不是模拟结果）
 *   - createFreshAdapter()（重新从同一持久化位置读状态）
 *   - corruptState(sessionId)（仅测试目录内，令该会话的状态不可读取）
 *
 * 此测试刻意将 UI 操作留在适配器外：目录选择应由浏览器 UI 触发，但路径判断、
 * 状态持久化及 hook 阻断必须可以脱离 DOM 用真实生产函数验证。
 */
import assert from 'node:assert/strict';
import { mkdir, mkdtemp, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import { pathToFileURL } from 'node:url';
import { parseHookOutput } from '@deepseek-ai/dsh-hook-protocol';

const workspaceHome = await mkdtemp(join(tmpdir(), 'dsh-session-office-gate-'));
const originalDshHome = process.env.DSH_HOME;
process.env.DSH_HOME = workspaceHome;
const adapterArg = process.argv[2] ?? './tests/session-office-gate.adapter.mjs';
const adapterUrl = adapterArg.includes(':') ? adapterArg : pathToFileURL(resolve(adapterArg)).href;
const { createSessionOfficeGateTestAdapter } = await import(adapterUrl);
assert.equal(typeof createSessionOfficeGateTestAdapter, 'function',
  '适配器必须导出 createSessionOfficeGateTestAdapter({ workspaceHome })');
let passed = 0;
let failed = 0;

async function test(name, body) {
  try {
    await body();
    passed += 1;
    console.log(`[ok] ${name}`);
  } catch (error) {
    failed += 1;
    console.error(`[FAIL] ${name}: ${error instanceof Error ? error.message : String(error)}`);
  }
}

function expectDenied(result, name, reasonPattern = /会话目录|办公工作区|确认|选择/i) {
  assert.equal(result && typeof result, 'object', `${name} 必须返回 hook 子进程结果`);
  const parsed = parseHookOutput(result.exitCode, result.stdout ?? '', result.stderr ?? '', 'UserPromptSubmit');
  const denied = result.exitCode === 2 || parsed.decision === 'deny' || parsed.decision === 'block';
  assert.equal(denied, true, `${name} 必须由 exit 2 或 UserPromptSubmit deny 阻断`);
  assert.match(String(parsed.reason ?? result.stderr ?? result.stdout ?? ''), reasonPattern,
    `${name} 的阻断结果必须向使用者说明缺少哪一步`);
}

function expectAllowed(result, name) {
  const parsed = parseHookOutput(result.exitCode, result.stdout ?? '', result.stderr ?? '', 'UserPromptSubmit');
  assert.notEqual(result.exitCode, 2, `${name} 不应以 exit 2 阻断`);
  assert.notEqual(parsed.decision, 'deny', `${name} 不应返回 UserPromptSubmit deny`);
  assert.notEqual(parsed.decision, 'block', `${name} 不应返回 block`);
}

function additionalContext(result, name) {
  assert.equal(result.exitCode, 0, `${name} hook 子进程必须正常结束`);
  const output = JSON.parse(result.stdout);
  assert.equal(output?.hookSpecificOutput?.hookEventName, 'UserPromptSubmit');
  assert.equal(typeof output?.hookSpecificOutput?.additionalContext, 'string',
    `${name} 必须输出模型可见的 additionalContext`);
  return output.hookSpecificOutput.additionalContext;
}

try {
  let gate = await createSessionOfficeGateTestAdapter({ workspaceHome });
  const fixtures = join(workspaceHome, 'fixtures');
  const root = join(fixtures, 'session-root');
  const child = join(root, 'office');
  const prefixSibling = join(fixtures, 'session-root-copy');
  const escaped = join(fixtures, 'outside');
  const otherDrive = 'E:\\DSH\\contract-fixtures\\office';
  const sessionA = 'contract-session-a';
  const sessionB = 'contract-session-b';

  await mkdir(child, { recursive: true });
  await mkdir(prefixSibling, { recursive: true });
  await mkdir(escaped, { recursive: true });

  await test('未选择任何目录时普通聊天被服务端阻断', async () => {
    expectDenied(await gate.runUserPromptHook(sessionA, '你好'), '未选择会话目录');
  });

  await test('仅选择会话目录后普通聊天仍被服务端阻断', async () => {
    await gate.selectSessionDirectory(sessionA, root);
    expectDenied(await gate.runUserPromptHook(sessionA, '继续'), '未选择办公工作区');
  });

  await test('办公工作区等于会话目录时放行', async () => {
    await gate.selectOfficeWorkspace(sessionA, root);
    expectAllowed(await gate.runUserPromptHook(sessionA, '现在可以聊天'), '相同目录');
  });

  await test('真正子目录可以作为办公工作区', async () => {
    await gate.selectSessionDirectory(sessionB, root);
    await gate.selectOfficeWorkspace(sessionB, child);
    const result = await gate.runUserPromptHook(sessionB, '子目录已确认');
    expectAllowed(result, '真正子目录');
    const context = additionalContext(result, '真正子目录');
    assert.match(context, new RegExp(JSON.stringify(root).replace(/[.*+?^${}()|[\]\\]/g, '\\$&')),
      '已确认上下文必须携带会话目录的 JSON 数据');
    assert.match(context, new RegExp(JSON.stringify(child).replace(/[.*+?^${}()|[\]\\]/g, '\\$&')),
      '已确认上下文必须携带办公子目录的 JSON 数据');
    assert.match(context, /无需再次询问/, '已确认上下文必须阻止重复目录提问');
  });

  await test('同前缀兄弟目录不能作为子目录接受', async () => {
    const id = 'contract-prefix-sibling';
    await gate.selectSessionDirectory(id, root);
    await assert.rejects(() => gate.selectOfficeWorkspace(id, prefixSibling), /子目录|范围|workspace|path/i);
  });

  await test('含 .. 的逃逸路径不能作为子目录接受', async () => {
    const id = 'contract-dotdot-escape';
    await gate.selectSessionDirectory(id, root);
    await assert.rejects(() => gate.selectOfficeWorkspace(id, escaped), /子目录|范围|workspace|path/i);
  });

  await test('不同盘符不能作为子目录接受', async () => {
    assert.equal(gate.sameOrChild(root, otherDrive), false,
      'Windows 不同盘符必须不是会话目录的子目录');
  });

  await test('不同会话不能借用已完成会话的确认', async () => {
    const id = 'contract-isolated';
    expectDenied(await gate.runUserPromptHook(id, '不应借用 session A'), '会话 ID 隔离');
  });

  await test('旧会话兼容放行但不假称网页已确认目录', async () => {
    const legacyId = 'contract-legacy-session';
    await gate.markLegacy(legacyId);
    await gate.selectSessionDirectory(legacyId, root);
    const result = await gate.runUserPromptHook(legacyId, '旧会话继续');
    expectAllowed(result, '旧会话兼容');
    const context = additionalContext(result, '旧会话兼容');
    assert.doesNotMatch(context, /本会话已由网页确认的事实/, '旧会话不得伪造网页确认事实');
    assert.doesNotMatch(context, new RegExp(JSON.stringify(root).replace(/[.*+?^${}()|[\]\\]/g, '\\$&')),
      '旧会话上下文不得注入未验证的会话目录');
  });

  await test('重新加载后已完成会话恢复，未完成会话仍阻断', async () => {
    gate = await gate.createFreshAdapter();
    expectAllowed(await gate.runUserPromptHook(sessionA, '刷新后已完成会话'), '刷新后的已完成会话');
    expectDenied(await gate.runUserPromptHook('contract-isolated', '刷新后未完成会话'), '刷新后的未完成会话');
  });

  await test('状态读取失败时拒绝并给出明确提示', async () => {
    await gate.corruptState(sessionA);
    const result = await gate.runUserPromptHook(sessionA, '状态已损坏');
    expectDenied(result, '状态读取失败', /状态|读取|恢复|重新选择/i);
    const parsed = parseHookOutput(result.exitCode, result.stdout ?? '', result.stderr ?? '', 'UserPromptSubmit');
    assert.match(String(parsed.reason ?? result.stderr ?? result.stdout ?? ''), /状态|读取|恢复|重新选择/i,
      '状态读取失败必须说明恢复动作，不能静默放行');
  });
} finally {
  if (originalDshHome === undefined) delete process.env.DSH_HOME;
  else process.env.DSH_HOME = originalDshHome;
  await rm(workspaceHome, { recursive: true, force: true });
}

console.log(`RESULT: ${passed} passed, ${failed} failed`);
process.exitCode = failed === 0 ? 0 : 1;
