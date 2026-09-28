const test = require('node:test');
const assert = require('node:assert/strict');
const { EventEmitter } = require('node:events');
const http = require('node:http');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { ensureLocalModel, readQwen4Config } = require('../local-qwen-autostart.cjs');

function logs() { return { info() {}, stdout() {}, stderr() {} }; }
function fixture(t) {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), 'dsh-qwen4-'));
  const server = path.join(root, 'llama-server.exe');
  const model = path.join(root, 'Qwen3-4B-Instruct-2507-Q8_0.gguf');
  fs.writeFileSync(server, 'fixture'); fs.writeFileSync(model, 'fixture');
  t.after(() => fs.rmSync(root, { recursive: true, force: true }));
  return { root, server, model };
}
async function modelServer(alias, key) {
  const server = http.createServer((req, res) => {
    if (req.headers.authorization !== `Bearer ${key}`) { res.writeHead(401); return res.end(); }
    res.setHeader('content-type', 'application/json'); res.end(JSON.stringify({ data: [{ id: alias }] }));
  });
  await new Promise(resolve => server.listen(0, '127.0.0.1', resolve));
  return server;
}
function fakeChild() {
  const child = new EventEmitter(); child.exitCode = null; child.killed = false;
  child.stdout = new EventEmitter(); child.stderr = new EventEmitter();
  child.kill = () => { child.killed = true; child.exitCode = 0; child.emit('close', 0); return true; };
  return child;
}
function env(files, port, extra = {}) { return { LOCAL_QWEN4_AUTOSTART: '1', LOCAL_QWEN4_SERVER: files.server,
  LOCAL_QWEN4_MODEL_PATH: files.model, LOCAL_QWEN4_PORT: String(port), LOCAL_QWEN_API_KEY: 'test-key', ...extra }; }

test('uses the separate Qwen 4B environment names and approved CPU defaults', t => {
  const files = fixture(t); const config = readQwen4Config(env(files, 18082));
  assert.equal(config.port, 18082); assert.equal(config.context, 32768); assert.equal(config.threads, 6);
  assert.equal(config.alias, 'local-qwen3-4b-instruct-2507');
});

test('reuses a matching local service and never stops it', async t => {
  const files = fixture(t); const server = await modelServer('local-qwen3-4b-instruct-2507', 'test-key');
  t.after(() => server.close()); const port = server.address().port;
  let spawned = false;
  const result = await ensureLocalModel(files.root, env(files, port), logs(), { spawn: () => { spawned = true; } });
  assert.equal(result.owned, false); await result.stop(); assert.equal(spawned, false); assert.equal(server.listening, true);
});

test('starts and stops only its own child after the model becomes healthy', async t => {
  const files = fixture(t); const holder = http.createServer();
  await new Promise(resolve => holder.listen(0, '127.0.0.1', resolve)); const port = holder.address().port;
  await new Promise(resolve => holder.close(resolve));
  let child; let started; let loading = true;
  t.after(() => started && started.close());
  const result = await ensureLocalModel(files.root, env(files, port), logs(), {
    spawn: () => {
      child = fakeChild();
      started = http.createServer((req, res) => {
        assert.equal(req.headers.authorization, 'Bearer test-key');
        if (loading) { res.writeHead(503); return res.end(); }
        res.setHeader('content-type', 'application/json'); res.end(JSON.stringify({ data: [{ id: 'local-qwen3-4b-instruct-2507' }] }));
      });
      started.listen(port, '127.0.0.1');
      setTimeout(() => { loading = false; }, 10);
      return child;
    }, pollIntervalMs: 1, startupTimeoutMs: 100,
  });
  assert.equal(result.owned, true); await result.stop(); assert.equal(child.killed, true);
});

test('reports a non-matching service as a port conflict without spawning', async t => {
  const files = fixture(t); const server = await modelServer('another-model', 'test-key'); t.after(() => server.close());
  await assert.rejects(ensureLocalModel(files.root, env(files, server.address().port), logs(), { spawn: () => assert.fail('must not spawn') }), /端口 .*其他服务占用/);
});

test('treats a pre-existing loading response as an external port conflict', async t => {
  const files = fixture(t); const server = http.createServer((_, res) => { res.writeHead(503); res.end(); });
  await new Promise(resolve => server.listen(0, '127.0.0.1', resolve)); t.after(() => server.close());
  await assert.rejects(ensureLocalModel(files.root, env(files, server.address().port), logs(), {
    spawn: () => assert.fail('must not spawn over an existing service'),
  }), /端口 .*HTTP 503/);
});

test('cleans up an owned child when startup fails', async t => {
  const files = fixture(t); const child = fakeChild();
  const result = ensureLocalModel(files.root, env(files, 65530), logs(), { spawn: () => { process.nextTick(() => child.emit('error', new Error('cannot load model'))); return child; }, pollIntervalMs: 1, startupTimeoutMs: 30 });
  await assert.rejects(result, /启动失败.*cannot load model/); assert.equal(child.killed, true);
});

test('reports a signal-based child exit while loading', async t => {
  const files = fixture(t); const child = fakeChild();
  const result = ensureLocalModel(files.root, env(files, 65528), logs(), {
    spawn: () => { process.nextTick(() => child.emit('close', null, 'SIGTERM')); return child; },
    pollIntervalMs: 1, startupTimeoutMs: 30,
  });
  await assert.rejects(result, /信号 SIGTERM/);
});

test('times out while a child-owned model remains unavailable and stops that child', async t => {
  const files = fixture(t); const child = fakeChild();
  await assert.rejects(ensureLocalModel(files.root, env(files, 65529), logs(), {
    spawn: () => child, pollIntervalMs: 1, startupTimeoutMs: 15,
  }), /15ms 内未就绪/);
  assert.equal(child.killed, true);
});
