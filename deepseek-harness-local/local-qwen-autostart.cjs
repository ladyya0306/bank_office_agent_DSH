// Starts only the approved local Qwen 4B server when DSH explicitly requests it.
const http = require('node:http');
const { spawn } = require('node:child_process');
const { readConfig, QWEN4_DEFAULTS, buildServerArgs } = require('./local-qwen.cjs');

const QWEN4_NAMES = Object.freeze({
  serverPath: 'LOCAL_QWEN4_SERVER',
  modelPath: 'LOCAL_QWEN4_MODEL_PATH',
  port: 'LOCAL_QWEN4_PORT',
  context: 'LOCAL_QWEN4_CONTEXT',
  threads: 'LOCAL_QWEN4_THREADS',
  apiKey: 'LOCAL_QWEN_API_KEY',
});
const NOOP = Object.freeze({ owned: false, stop: async () => {} });
const closedChildren = new WeakSet();

function readQwen4Config(env) {
  return readConfig(env, { defaults: QWEN4_DEFAULTS, names: QWEN4_NAMES });
}

function modelProbe(config, request = http.request) {
  return new Promise(resolve => {
    const req = request({ host: '127.0.0.1', port: config.port, path: '/v1/models', method: 'GET',
      headers: { Authorization: `Bearer ${config.apiKey}` }, timeout: 1500 }, response => {
      let body = '';
      response.setEncoding('utf8');
      response.on('data', chunk => { body += chunk; });
      response.on('end', () => {
        if (response.statusCode === 503) return resolve({ kind: 'loading', detail: 'HTTP 503' });
        if (response.statusCode !== 200) return resolve({ kind: 'occupied', detail: `HTTP ${response.statusCode}` });
        try {
          const data = JSON.parse(body).data;
          const matched = Array.isArray(data) && data.some(model => model && model.id === config.alias);
          resolve(matched ? { kind: 'matched' } : { kind: 'occupied', detail: '接口没有声明所需模型别名' });
        } catch {
          resolve({ kind: 'occupied', detail: '接口返回的模型清单无效' });
        }
      });
    });
    req.on('timeout', () => req.destroy(new Error('timeout')));
    req.on('error', error => {
      if (error.code === 'ECONNREFUSED') resolve({ kind: 'absent' });
      else resolve({ kind: 'unreachable', detail: error.code || error.message });
    });
    req.end();
  });
}

function portError(config, probe) {
  return new Error(`本地 Qwen 4B 端口 ${config.port} 已被其他服务占用：${probe.detail}。为避免停止外部服务，DSH未启动模型。`);
}

function wait(ms) { return new Promise(resolve => setTimeout(resolve, ms)); }

function trackChildClose(child) {
  child.once('close', () => closedChildren.add(child));
}

function waitForExit(child, timeoutMs) {
  if (closedChildren.has(child)) return Promise.resolve();
  return new Promise(resolve => {
    const timer = setTimeout(done, timeoutMs);
    function done() { clearTimeout(timer); child.removeListener('close', done); resolve(); }
    child.once('close', done);
  });
}

async function stopOwnedChild(child, timeoutMs) {
  if (closedChildren.has(child)) return;
  child.kill('SIGINT');
  await waitForExit(child, timeoutMs);
  if (!closedChildren.has(child)) {
    child.kill('SIGTERM');
    await waitForExit(child, timeoutMs);
  }
  if (!closedChildren.has(child)) {
    child.kill('SIGKILL');
    await waitForExit(child, timeoutMs);
  }
}

function pipeChildLogs(child, logs) {
  const stdout = chunk => logs.stdout(chunk);
  const stderr = chunk => logs.stderr(chunk);
  child.stdout.on('data', stdout);
  child.stderr.on('data', stderr);
  return () => {
    child.stdout.removeListener('data', stdout);
    child.stderr.removeListener('data', stderr);
  };
}

async function waitForModel(config, child, options) {
  const deadline = Date.now() + options.startupTimeoutMs;
  let childFailure;
  const failed = error => { childFailure = error || new Error('模型服务提前退出'); };
  const closed = (code, signal) => failed(new Error(signal
    ? `模型服务提前被信号 ${signal} 停止`
    : `模型服务提前退出（退出码 ${code}）`));
  child.once('error', failed);
  child.once('close', closed);
  try {
    while (Date.now() <= deadline) {
      if (childFailure) throw new Error(`本地 Qwen 4B 启动失败：${childFailure.message}`);
      const probe = await modelProbe(config, options.request);
      if (probe.kind === 'matched') return;
      if (probe.kind === 'loading') {
        await options.wait(options.pollIntervalMs);
        continue;
      }
      if (probe.kind === 'occupied') throw portError(config, probe);
      if (probe.kind === 'unreachable') throw new Error(`无法检查本地 Qwen 4B 端口 ${config.port}：${probe.detail}`);
      await options.wait(options.pollIntervalMs);
    }
    throw new Error(`本地 Qwen 4B 在 ${options.startupTimeoutMs}ms 内未就绪。`);
  } finally {
    child.removeListener('error', failed);
    child.removeListener('close', closed);
  }
}

async function ensureLocalModel(harnessRoot, env = process.env, logs, injected = {}) {
  if (env.LOCAL_QWEN4_AUTOSTART !== '1') return NOOP;
  const config = readQwen4Config(env);
  if (!config.apiKey || config.apiKey === 'local-only') throw new Error('LOCAL_QWEN_API_KEY 未设置；无法安全连接本地 Qwen 4B 服务。');
  const options = {
    spawn: injected.spawn || spawn, request: injected.request || http.request, wait: injected.wait || wait,
    startupTimeoutMs: injected.startupTimeoutMs || 90000, pollIntervalMs: injected.pollIntervalMs || 1000,
    stopTimeoutMs: injected.stopTimeoutMs || 5000,
  };
  const before = await modelProbe(config, options.request);
  if (before.kind === 'matched') {
    logs.info(`复用已就绪的本地 Qwen 4B 服务：127.0.0.1:${config.port}。`);
    return NOOP;
  }
  if (before.kind === 'occupied' || before.kind === 'loading') throw portError(config, before);
  if (before.kind === 'unreachable') throw new Error(`无法检查本地 Qwen 4B 端口 ${config.port}：${before.detail}`);
  const child = options.spawn(config.serverPath, buildServerArgs(config), {
    shell: false, windowsHide: true, stdio: ['ignore', 'pipe', 'pipe'],
    env: { ...env, LLAMA_API_KEY: config.apiKey },
  });
  trackChildClose(child);
  const detachLogs = pipeChildLogs(child, logs);
  logs.info(`正在启动本地 Qwen 4B：127.0.0.1:${config.port}。`);
  try {
    await waitForModel(config, child, options);
  } catch (error) {
    await stopOwnedChild(child, options.stopTimeoutMs);
    detachLogs();
    throw error;
  }
  let stopped = false;
  return { owned: true, async stop() {
    if (stopped) return;
    stopped = true;
    await stopOwnedChild(child, options.stopTimeoutMs);
    detachLogs();
  } };
}

module.exports = { QWEN4_NAMES, readQwen4Config, modelProbe, ensureLocalModel };
