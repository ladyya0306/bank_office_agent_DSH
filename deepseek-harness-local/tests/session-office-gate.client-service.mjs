/**
 * Loads the production browser module with a minimal Cordis/React surface.
 * It tests the public uiWorkspace.startSession interception without a browser.
 */
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import vm from 'node:vm';

const clientPath = new URL('../workspace/plugins/session-workspace-gate/lib/client.js', import.meta.url);
let descriptor;
const source = await readFile(clientPath, 'utf8');
vm.runInNewContext(source, {
  console,
  window: { __ModuleLoader__: { load: (value) => { descriptor = value; } } },
}, { filename: clientPath.pathname });
assert.ok(descriptor, '生产 client.js 必须注册 ModuleLoader 模块');

const react = {
  useState: () => { throw new Error('本测试不渲染 GateDock'); },
  useEffect: () => { throw new Error('本测试不渲染 GateDock'); },
  useSyncExternalStore: () => { throw new Error('本测试不渲染 overlay'); },
};
const jsx = { jsx: (type, props) => ({ type, props }), jsxs: (type, props) => ({ type, props }) };
const exported = descriptor.factory((name) => {
  if (name === 'react') return react;
  if (name === 'react/jsx-runtime') return jsx;
  throw new Error(`unexpected production client dependency: ${name}`);
});
assert.equal(typeof exported.apply, 'function', '生产 client.js 必须导出 apply');

function deferred() {
  let resolve;
  let reject;
  const promise = new Promise((res, rej) => { resolve = res; reject = rej; });
  return { promise, resolve, reject };
}

async function eventually(check, message) {
  for (let i = 0; i < 20; i += 1) {
    if (check()) return;
    await new Promise((resolve) => setImmediate(resolve));
  }
  assert.fail(message);
}

function setup({ picker }) {
  const calls = { clear: 0, picker: 0, create: [], freshSessions: [], opened: [], originalStart: [], registrations: [] };
  const sessions = {
    clear: () => { calls.clear += 1; },
    create: async ({ workspaceId }) => {
      const id = `session:${workspaceId}:${calls.freshSessions.length + 1}`;
      calls.freshSessions.push({ workspaceId, id });
      return id;
    },
    list: {
      getSnapshot: () => ({ current: undefined, byId: {} }),
      subscribe: () => () => {},
    },
  };
  const workspaces = {
    create: async ({ path }) => {
      calls.create.push(path);
      return { workspaceId: `workspace:${path}`, path };
    },
    list: { getSnapshot: () => ({ items: [] }), subscribe: () => () => {} },
  };
  const uiWorkspace = {
    startSession: (workspaceId) => { calls.originalStart.push(workspaceId); },
    openSession: (sessionId) => { calls.opened.push(sessionId); },
    pickDirectory: () => { calls.picker += 1; return picker(); },
  };
  const blockValues = new Map();
  const blocks = {
    set: (sessionId, block) => { blockValues.set(sessionId, block); },
    storeFor: (sessionId) => ({ getSnapshot: () => blockValues.get(sessionId) }),
  };
  const services = new Map([
    ['sessions', sessions], ['workspaces', workspaces], ['conversation', { blocks }], ['uiWorkspace', uiWorkspace],
  ]);
  const ctx = {
    get: (name) => services.get(name),
    effect: (effect) => effect(),
    slots: {
      inject: (_name, register) => { calls.registrations.push(register()); },
      register: (definition, component) => ({ definition, component }),
    },
  };
  exported.apply(ctx);
  const gateRegistration = calls.registrations.find(({ definition }) => definition.id === 'session-workspace-gate');
  assert.ok(gateRegistration, '生产插件必须注册 composer gate');
  const gateBlocks = gateRegistration.component({}).props.blocks;
  return { calls, uiWorkspace, blocks, gateBlocks };
}

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

await test('点击新会话先清空并打开目录选择器', async () => {
  const pick = deferred();
  const { calls, uiWorkspace } = setup({ picker: () => pick.promise });
  uiWorkspace.startSession();
  assert.equal(calls.clear, 1, '新会话开始时必须清空当前 session 选择');
  assert.equal(calls.picker, 1, '新会话开始时必须先打开一次目录选择器');
  assert.deepEqual(calls.create, [], '目录未选定前不得创建 workspace');
  assert.deepEqual(calls.freshSessions, [], '目录未选定前不得创建新 session');
  pick.resolve(null);
  await eventually(() => true, '取消路径应能收束');
  assert.deepEqual(calls.freshSessions, [], '取消后不得创建新 session');
  assert.deepEqual(calls.opened, [], '取消后不得打开或发送到任何 session');
  assert.deepEqual(calls.originalStart, [], '取消后不得回退调用原 startSession');
});

await test('选择目录后创建 workspace、独立 session 并打开该 session', async () => {
  const pick = deferred();
  const { calls, uiWorkspace } = setup({ picker: () => pick.promise });
  uiWorkspace.startSession();
  pick.resolve('D:\\contract-fixtures\\chosen-session');
  await eventually(() => calls.opened.length === 1, '选择路径后应打开新 session');
  assert.deepEqual(calls.create, ['D:\\contract-fixtures\\chosen-session']);
  assert.deepEqual(calls.freshSessions, [{
    workspaceId: 'workspace:D:\\contract-fixtures\\chosen-session',
    id: 'session:workspace:D:\\contract-fixtures\\chosen-session:1',
  }]);
  assert.deepEqual(calls.opened, ['session:workspace:D:\\contract-fixtures\\chosen-session:1']);
  assert.deepEqual(calls.originalStart, [], '不得复用原 startSession 的空白 session 路径');
});

await test('同一目录完成首轮后再次新会话获得不同 session ID', async () => {
  const path = 'D:\\contract-fixtures\\same-directory';
  const { calls, uiWorkspace } = setup({ picker: () => Promise.resolve(path) });
  uiWorkspace.startSession();
  await eventually(() => calls.opened.length === 1, '首轮应打开 session');
  uiWorkspace.startSession();
  await eventually(() => calls.opened.length === 2, '第二轮应打开 session');
  assert.equal(calls.freshSessions.length, 2);
  assert.notEqual(calls.freshSessions[0].id, calls.freshSessions[1].id,
    '同目录的两次新会话必须使用不同 session ID');
  assert.deepEqual(calls.originalStart, [], '两轮都不得调用可能复用空白 session 的原 startSession');
});

await test('目录选择未完成时双击新会话不重复打开选择器', async () => {
  const pick = deferred();
  const { calls, uiWorkspace } = setup({ picker: () => pick.promise });
  uiWorkspace.startSession();
  uiWorkspace.startSession();
  assert.equal(calls.clear, 1, '重复点击不能重复清空');
  assert.equal(calls.picker, 1, '重复点击不能重复打开目录选择器');
  pick.resolve(null);
  await eventually(() => true, '取消路径应能收束');
});

await test('门禁未确认时外部插件清除 block 不能解除门禁', async () => {
  const { blocks, gateBlocks } = setup({ picker: () => Promise.resolve(null) });
  const sessionId = 'session-gate-priority';
  gateBlocks.set(sessionId, { reason: '门禁尚未确认' });
  blocks.set(sessionId, undefined);
  assert.deepEqual(blocks.storeFor(sessionId).getSnapshot(), { reason: '门禁尚未确认' });
});

await test('门禁解除后恢复外部模型选择 block', async () => {
  const { blocks, gateBlocks } = setup({ picker: () => Promise.resolve(null) });
  const sessionId = 'session-model-selection';
  const modelBlock = { reason: '请选择可用模型' };
  gateBlocks.set(sessionId, { reason: '门禁尚未确认' });
  blocks.set(sessionId, modelBlock);
  assert.deepEqual(blocks.storeFor(sessionId).getSnapshot(), { reason: '门禁尚未确认' },
    '门禁生效期间应保持门禁 block');
  gateBlocks.set(sessionId, undefined);
  assert.deepEqual(blocks.storeFor(sessionId).getSnapshot(), modelBlock,
    '解除门禁后必须恢复外部模型选择 block');
});

console.log(`RESULT: ${passed} passed, ${failed} failed`);
process.exitCode = failed === 0 ? 0 : 1;
