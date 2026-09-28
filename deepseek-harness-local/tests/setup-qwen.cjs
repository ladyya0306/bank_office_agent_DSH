const { test } = require('node:test');
const assert = require('node:assert/strict');
const YAML = require('yaml');
const { profiles, updateSettings } = require('../setup-qwen.cjs');
const { ensureLocalKey } = require('../setup-qwen.cjs');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');

test('generates local authentication without changing other keys and preserves it on rerun', t => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'qwen-key-'));
  t.after(() => fs.rmSync(dir, { recursive: true, force: true }));
  const file = path.join(dir, '.env');
  fs.writeFileSync(file, 'DEEPSEEK_API_KEY=keep-test-value\nLOCAL_QWEN_API_KEY=local-only\n');
  const env = { LOCAL_QWEN_API_KEY: 'local-only' };
  ensureLocalKey(dir, env);
  assert.match(env.LOCAL_QWEN_API_KEY, /^[a-f0-9]{64}$/);
  const before = fs.readFileSync(file, 'utf8');
  assert.ok(before.includes('DEEPSEEK_API_KEY=keep-test-value\n'));
  ensureLocalKey(dir, env);
  assert.equal(fs.readFileSync(file, 'utf8'), before);
});

test('adds routes while preserving Flash default, existing providers and permissions', () => {
  const old = { 'agent-default-model': { provider: 'deepseek', model: 'deepseek-flash' },
    permission: { preset: 'workspace-write' }, 'llm-pi-ai': { providers: { existing: { apiKeyEnv: 'EXISTING_KEY' } } } };
  const text = updateSettings(YAML.stringify(old), profiles({ DASHSCOPE_API_KEY: 'SECRET-DO-NOT-SAVE' }));
  const next = YAML.parse(text);
  assert.deepEqual(next['agent-default-model'], old['agent-default-model']);
  assert.deepEqual(next.permission, old.permission);
  assert.deepEqual(next['llm-pi-ai'].providers.existing, old['llm-pi-ai'].providers.existing);
  assert.equal(text.includes('SECRET-DO-NOT-SAVE'), false);
  assert.equal(updateSettings(text, profiles({})), text);
});
test('local provider uses loopback, bounded context, and no advertised thinking effort', () => {
  const local = profiles({})['local-qwen'];
  assert.equal(local.baseURL, 'http://127.0.0.1:18081/v1');
  assert.equal(local.models[0].contextWindow, 32768);
  assert.equal(local.models[0].reasoningEfforts, false);
  assert.equal(local.compat.chatTemplateKwargs.enable_thinking, false);
});
test('invalid configuration fails instead of writing a broken route', () => {
  assert.throws(() => profiles({ LOCAL_QWEN_PORT: 'NaN' }));
  assert.throws(() => profiles({ LOCAL_QWEN_CONTEXT: '999999' }));
  assert.throws(() => profiles({ DASHSCOPE_BASE_URL: 'http://example.com' }));
  assert.throws(() => updateSettings('wrong: [', profiles({})));
});
