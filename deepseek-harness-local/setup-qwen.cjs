// Register extra routes using DSH's existing settings adapter; never copy credentials.
const fs = require('node:fs');
const path = require('node:path');
const YAML = require('yaml');
const { randomBytes } = require('node:crypto');

function ensureLocalKey(root, env) {
  if (env.LOCAL_QWEN_API_KEY && env.LOCAL_QWEN_API_KEY !== 'local-only') return;
  const file = path.join(root, '.env');
  const original = fs.existsSync(file) ? fs.readFileSync(file, 'utf8') : '';
  const key = randomBytes(32).toString('hex');
  const line = `LOCAL_QWEN_API_KEY=${key}`;
  const pattern = /^LOCAL_QWEN_API_KEY=.*$/m;
  const next = pattern.test(original) ? original.replace(pattern, line) : `${original}\n${line}\n`;
  fs.writeFileSync(file, next, { encoding: 'utf8', mode: 0o600 });
  env.LOCAL_QWEN_API_KEY = key;
}

function positive(value, fallback, max, label) {
  const n = value === undefined || value === '' ? fallback : Number(value);
  if (!Number.isInteger(n) || n < 1 || n > max) throw new Error(`${label} 超出有效范围`);
  return n;
}

function profiles(env) {
  const port = positive(env.LOCAL_QWEN_PORT, 18081, 65535, '本地端口');
  const context = positive(env.LOCAL_QWEN_CONTEXT, 32768, 32768, '本地上下文');
  if (context < 4096) throw new Error('本地上下文至少4096；DSH工具说明需要空间');
  const port4 = positive(env.LOCAL_QWEN4_PORT, 18082, 65535, '4B端口');
  const context4 = positive(env.LOCAL_QWEN4_CONTEXT, 32768, 32768, '4B上下文');
  if (context4 < 4096) throw new Error('4B上下文至少4096');
  const baseURL = env.DASHSCOPE_BASE_URL || 'https://dashscope.aliyuncs.com/compatible-mode/v1';
  const url = new URL(baseURL);
  if (url.protocol !== 'https:' || url.username || url.password || url.search || url.hash) {
    throw new Error('百炼地址须为不含凭据或查询参数的HTTPS接口地址');
  }
  const common = { api: 'openai-completions', retryPolicy: { mode: 'normal', maxRetries: 1 } };
  const compat = { supportsStore: false, supportsDeveloperRole: false,
    supportsReasoningEffort: false, supportsStrictMode: false, maxTokensField: 'max_tokens' };
  return {
    'bailian-qwen': { ...common, displayName: '阿里云百炼（北京）', apiKeyEnv: 'DASHSCOPE_API_KEY', baseURL,
      compat: { ...compat, thinkingFormat: 'qwen' },
      models: [{ id: env.DASHSCOPE_MODEL || 'qwen-plus', name: 'Qwen Plus · 百炼',
        contextWindow: 131072, maxTokens: 8192, input: ['text'], reasoningEfforts: false }] },
    'local-qwen': { ...common, displayName: '本机 Qwen · 离线', apiKeyEnv: 'LOCAL_QWEN_API_KEY',
      baseURL: `http://127.0.0.1:${port}/v1`, timeoutMs: 600000, streamIdleTimeoutMs: 300000,
      compat: { ...compat, thinkingFormat: 'chat-template', chatTemplateKwargs: { enable_thinking: false } },
      models: [{ id: 'local-qwen3-1.7b', name: 'Qwen3-1.7B Q8_0 · 本地实验',
        contextWindow: context, maxTokens: 2048, input: ['text'], reasoningEfforts: false }] },
    'local-qwen4': { ...common, displayName: '本机 Qwen 4B · 离线', apiKeyEnv: 'LOCAL_QWEN_API_KEY',
      baseURL: `http://127.0.0.1:${port4}/v1`, timeoutMs: 600000, streamIdleTimeoutMs: 300000,
      compat: { ...compat, thinkingFormat: 'chat-template', chatTemplateKwargs: { enable_thinking: false } },
      models: [{ id: 'local-qwen3-4b-instruct-2507', name: 'Qwen3-4B Instruct 2507 Q4_K_M · 本地',
        contextWindow: context4, maxTokens: 4096, input: ['text'], reasoningEfforts: false }] }
  };
}

function updateSettings(text, routes) {
  const doc = YAML.parseDocument(text || '{}');
  if (doc.errors.length) throw new Error('DSH settings.yaml 解析失败，未修改');
  if (!YAML.isMap(doc.contents)) throw new Error('DSH settings.yaml 顶层必须为映射');
  for (const [name, profile] of Object.entries(routes)) doc.setIn(['llm-pi-ai', 'providers', name], profile);
  return doc.toString();
}

function setup(root, env) {
  const file = path.join(root, 'workspace', 'settings.yaml');
  const before = fs.existsSync(file) ? fs.readFileSync(file, 'utf8') : '{}\n';
  const after = updateSettings(before, profiles(env));
  if (after === before) return { changed: false };
  const dir = path.join(root, 'workspace', 'backups', `qwen-models-${Date.now()}`);
  fs.mkdirSync(dir, { recursive: true });
  if (fs.existsSync(file)) fs.copyFileSync(file, path.join(dir, 'settings.yaml'));
  const temp = `${file}.${process.pid}.tmp`;
  fs.writeFileSync(temp, after, { encoding: 'utf8', flag: 'wx' });
  fs.renameSync(temp, file);
  return { changed: true, backup: dir };
}

if (require.main === module) {
  try {
    require('dotenv').config({ path: path.join(__dirname, '.env') });
    ensureLocalKey(__dirname, process.env);
    console.log(JSON.stringify(setup(__dirname, process.env)));
    console.log('模型选项已保存。填写DASHSCOPE_API_KEY后重启DSH；4B按LOCAL_QWEN4_AUTOSTART配置随DSH启动；1.7B仍用npm run local-qwen启动。');
  } catch (error) { console.error(error.message); process.exitCode = 1; }
}
module.exports = { profiles, updateSettings, setup, ensureLocalKey };
