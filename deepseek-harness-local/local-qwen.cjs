// Local-only launcher for the approved Qwen3 GGUF test model.
const fs = require('node:fs');
const path = require('node:path');
const { spawn } = require('node:child_process');
const { openRuntimeLogs } = require('./runtime-logging.cjs');

const DEFAULTS = Object.freeze({
  serverPath: 'D:\\MiniProj\\one-square-meter-civilization-vibe-starter\\build\\tools\\llama.cpp\\b10964\\llama-server.exe',
  modelPath: 'D:\\MiniProj\\one-square-meter-civilization-vibe-starter\\build\\models\\resident_planner\\Qwen3-1.7B-Q8_0.gguf',
  port: 18081,
  context: 32768,
  threads: 6,
  maxOutput: 4096,
  alias: 'local-qwen3-1.7b',
});

function integer(value, name, minimum, maximum) {
  if (!/^[0-9]+$/.test(String(value))) throw new Error(`${name} 必须是整数。`);
  const parsed = Number(value);
  if (!Number.isSafeInteger(parsed) || parsed < minimum || parsed > maximum) {
    throw new Error(`${name} 必须在 ${minimum} 到 ${maximum} 之间。`);
  }
  return parsed;
}

function existingFile(value, name, extension) {
  if (!value || typeof value !== 'string') throw new Error(`${name} 未设置。`);
  const resolved = path.resolve(value);
  if (extension && path.extname(resolved).toLowerCase() !== extension) {
    throw new Error(`${name} 必须是 ${extension} 文件。`);
  }
  let stat;
  try { stat = fs.statSync(resolved); } catch { throw new Error(`${name} 不存在：${resolved}`); }
  if (!stat.isFile()) throw new Error(`${name} 必须指向文件：${resolved}`);
  return resolved;
}

function readConfig(env = process.env) {
  return {
    serverPath: existingFile(env.LOCAL_QWEN_SERVER || DEFAULTS.serverPath, 'LOCAL_QWEN_SERVER'),
    modelPath: existingFile(env.LOCAL_QWEN_MODEL_PATH || DEFAULTS.modelPath, 'LOCAL_QWEN_MODEL_PATH', '.gguf'),
    port: integer(env.LOCAL_QWEN_PORT || DEFAULTS.port, 'LOCAL_QWEN_PORT', 1, 65535),
    context: integer(env.LOCAL_QWEN_CONTEXT || DEFAULTS.context, 'LOCAL_QWEN_CONTEXT', 1024, 32768),
    threads: integer(env.LOCAL_QWEN_THREADS || DEFAULTS.threads, 'LOCAL_QWEN_THREADS', 1, 256),
    maxOutput: DEFAULTS.maxOutput,
    alias: DEFAULTS.alias,
    apiKey: env.LOCAL_QWEN_API_KEY,
  };
}

function buildServerArgs(config) {
  return [
    '--model', config.modelPath,
    '--host', '127.0.0.1',
    '--port', String(config.port),
    '--ctx-size', String(config.context),
    '--threads', String(config.threads),
    '--parallel', '1',
    '--n-predict', String(config.maxOutput),
    '--device', 'none',
    '--n-gpu-layers', '0',
    '--alias', config.alias,
    '--jinja',
    '--reasoning', 'off',
    '--no-webui',
    '--cors-origins', 'http://127.0.0.1:3080',
  ];
}

function launch(config, root = __dirname) {
  if (!config.apiKey || config.apiKey === 'local-only') throw new Error('请先运行npm run setup-qwen，生成本地服务访问密钥。');
  const logs = openRuntimeLogs(root);
  const child = spawn(config.serverPath, buildServerArgs(config), {
    shell: false, stdio: ['ignore', 'pipe', 'pipe'], env: { ...process.env, LLAMA_API_KEY: config.apiKey },
  });
  child.stdout.on('data', logs.stdout);
  child.stderr.on('data', logs.stderr);
  child.on('error', error => {
    process.stderr.write(`本地 Qwen 服务未能启动：${error.message}\n`);
    logs.close();
    process.exitCode = 1;
  });
  child.on('close', code => {
    logs.close();
    process.exitCode = code || 0;
  });
  let stopping = false;
  const stop = signal => {
    if (stopping) return;
    stopping = true;
    process.stderr.write(`收到 ${signal}，正在停止本地 Qwen 服务。\n`);
    child.kill('SIGINT');
  };
  process.once('SIGINT', () => stop('Ctrl+C'));
  process.once('SIGTERM', () => stop('SIGTERM'));
  process.stdout.write(`正在启动本地 Qwen，等待模型加载；日志目录：${logs.directory}\n`);
  return child;
}

function main() {
  require('dotenv').config({ path: path.join(__dirname, '.env') });
  launch(readConfig());
}

if (require.main === module) main();

module.exports = { DEFAULTS, readConfig, buildServerArgs, launch };
