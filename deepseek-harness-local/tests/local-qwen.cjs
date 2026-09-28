const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { readConfig, buildServerArgs } = require('../local-qwen.cjs');

function fixture(t) {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), 'dsh-local-qwen-'));
  const server = path.join(root, 'llama-server.exe');
  const model = path.join(root, 'Qwen3-1.7B-Q8_0.gguf');
  fs.writeFileSync(server, 'fixture');
  fs.writeFileSync(model, 'fixture');
  t.after(() => fs.rmSync(root, { recursive: true, force: true }));
  return { server, model };
}

test('builds a loopback-only, CPU, single-slot Qwen command', t => {
  const { server, model } = fixture(t);
  const config = readConfig({ LOCAL_QWEN_SERVER: server, LOCAL_QWEN_MODEL_PATH: model });
  const args = buildServerArgs(config);
  assert.deepEqual(args.slice(0, 6), ['--model', model, '--host', '127.0.0.1', '--port', '18081']);
  assert.ok(args.includes('--device') && args[args.indexOf('--device') + 1] === 'none');
  assert.ok(args.includes('--parallel') && args[args.indexOf('--parallel') + 1] === '1');
  assert.ok(args.includes('--jinja'));
  assert.ok(args.includes('--cors-origins'));
  assert.ok(args.includes('--reasoning') && args[args.indexOf('--reasoning') + 1] === 'off');
});

test('rejects missing files and invalid context', t => {
  const { server, model } = fixture(t);
  assert.throws(() => readConfig({ LOCAL_QWEN_SERVER: server, LOCAL_QWEN_MODEL_PATH: path.join(path.dirname(model), 'missing.gguf') }), /不存在/);
  assert.throws(() => readConfig({ LOCAL_QWEN_SERVER: server, LOCAL_QWEN_MODEL_PATH: model, LOCAL_QWEN_CONTEXT: '0' }), /1024/);
  assert.throws(() => readConfig({ LOCAL_QWEN_SERVER: server, LOCAL_QWEN_MODEL_PATH: model, LOCAL_QWEN_CONTEXT: '32769' }), /32768/);
});

test('keeps hostile path text as one argument; no shell command is constructed', t => {
  const { server, model } = fixture(t);
  const hostile = `${model} & whoami`;
  const args = buildServerArgs({ ...readConfig({ LOCAL_QWEN_SERVER: server, LOCAL_QWEN_MODEL_PATH: model }), modelPath: hostile });
  assert.equal(args[args.indexOf('--model') + 1], hostile);
  assert.equal(args.filter(value => value.includes('& whoami')).length, 1);
  assert.equal(args.some(value => value === 'cmd.exe' || value === '/c'), false);
});
