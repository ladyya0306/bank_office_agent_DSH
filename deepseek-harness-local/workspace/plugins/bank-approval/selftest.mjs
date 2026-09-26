import assert from 'node:assert/strict';
import { Session } from '@deepseek-ai/dsh-session';
import { apply } from './index.mjs';

const listeners = new Map();
let policy = 'ask';
apply({
  on(name, listener) { listeners.set(name, listener); },
  get() { return { effectivePolicy: () => policy }; },
  logger: { warn(message) { throw new Error(message); } },
});

function session(cwd) {
  const events = [];
  return {
    header: { cwd },
    get seq() { return events.length; },
    eventAt(index) { return events[index]; },
    append(type, data) { events.push({ type, data }); },
    events,
  };
}
const pre = listeners.get('tools/pre-execute');
const approve = listeners.get('approval/request');
const first = session('D:\\office\\case-a');
const other = session('D:\\office\\case-b');
const tool = (s, callId, name, args) => ({ agent: { session: s }, callId, name, arguments: args });
const next = () => ({ kind: 'allow' });

assert.equal(pre(tool(first, 'a1', 'write', { path: 'D:\\office\\case-a\\one.txt' }), next).kind, 'ask');
assert.equal(await approve({ agent: { session: first }, callId: 'a1' }, () => 'allowed-always'), 'allowed-once');
assert.equal(pre(tool(first, 'a2', 'edit', { path: 'D:\\office\\case-a\\two.txt' }), next).kind, 'allow');
assert.equal(pre(tool(other, 'b1', 'write', { path: 'D:\\office\\case-b\\one.txt' }), next).kind, 'ask');
assert.equal(pre(tool(first, 'a3', 'pwsh', { command: 'python run.py' }), next).kind, 'ask');
assert.equal(pre(tool(first, 'a4', 'pwsh', { command: 'python -c "print(1)"' }), next).kind, 'deny');
assert.equal(pre(tool(first, 'a6', 'office_fill_task',
  { action: 'prepare', work: 'D:\\office\\case-a' }), next).kind, 'allow');
const unknown = pre(tool(first, 'a7', 'unlisted_custom_tool', { value: 'x' }), next);
assert.equal(unknown.kind, 'ask');
assert.match(unknown.reason, /没有登记过的工具/);
assert.equal(first.events[0].type, 'approval/policy');
assert.deepEqual(first.events[0].data.bankGrant,
  { scope: '写文件', cwd: first.header.cwd, callId: 'a1' });
policy = 'never';
assert.equal(pre(tool(first, 'a5', 'write', { path: 'D:\\office\\case-a\\three.txt' }), next).kind, 'deny');
assert.equal(pre(tool(first, 'a8', 'office_fill_task',
  { action: 'prepare', work: 'D:\\office\\case-a' }), next).kind, 'allow');
console.log('bank-approval: 允许一次 / 本会话同类一直允许 / 跨会话隔离 / 禁止清单 / never 策略通过');

// 真正的 DSH 会话记录要接受授权事件，并能从记录重建同会话授权。
policy = 'ask';
const persistedId = '00000000-0000-4000-8000-000000000001';
const persisted = Session.create(persistedId, undefined,
  { version: 3, id: persistedId, createdAt: Date.now(), isSeeded: false,
    cwd: 'D:\\office\\case-c' });
assert.equal(pre(tool(persisted, 'c1', 'write', { path: 'D:\\office\\case-c\\one.txt' }), next).kind, 'ask');
assert.equal(await approve({ agent: { session: persisted }, callId: 'c1' },
  () => 'allowed-always'), 'allowed-once');
assert.equal(pre(tool(persisted, 'c2', 'edit', { path: 'D:\\office\\case-c\\two.txt' }), next).kind, 'allow');
const resumed = Session.create(persistedId, persisted.snapshotEvents(),
  { version: 3, id: persistedId, createdAt: persisted.header.createdAt,
    isSeeded: false, cwd: persisted.header.cwd });
assert.equal(pre(tool(resumed, 'c3', 'write', { path: 'D:\\office\\case-c\\three.txt' }), next).kind, 'allow');
console.log('bank-approval: DSH 原生会话记录接受授权事件，恢复会话后仍有效');
