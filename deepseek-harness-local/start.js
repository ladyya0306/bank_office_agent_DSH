require('dotenv').config({ path: require('path').join(__dirname, '.env') });

const { spawn } = require('child_process');
const fs = require('fs');
const path = require('path');
const { prepareToolV2, portableWorkspace } = require('./tool-v2-startup.cjs');

const toolStatus = prepareToolV2(__dirname);
console.log(toolStatus.changed ? 'toolV2 接入已更新。' : 'toolV2 接入检查通过。');
process.env.WORKSPACE_DIR = portableWorkspace(__dirname, process.env.WORKSPACE_DIR);

const args = process.argv.slice(2);
const isDev = args.includes('--dev');
const sessionOfficeGate = args.includes('--session-office-gate');
const host = process.env.HOST || '127.0.0.1';
const port = process.env.PORT || '3080';
const localizedNativePicker = sessionOfficeGate && process.platform === 'win32' && host === '127.0.0.1';

const dshArgs = sessionOfficeGate ? ['--profile', 'web'] : ['web'];
if (sessionOfficeGate) {
  const workspace = path.resolve(process.env.WORKSPACE_DIR || path.join(__dirname, 'workspace'));
  const plugin = path.join(__dirname, 'workspace', 'plugins', 'session-workspace-gate', 'lib', 'index.js');
  const pickerPlugin = path.join(__dirname, 'workspace', 'plugins', 'session-directory-picker-zh', 'lib', 'index.js');
  const patchPath = path.join(workspace, 'profiles', 'web', 'session-office-gate.patch.yml');
  if (!fs.existsSync(plugin)) throw new Error(`Session office gate plugin missing: ${plugin}`);
  if (localizedNativePicker && !fs.existsSync(pickerPlugin)) throw new Error(`Localized directory picker plugin missing: ${pickerPlugin}`);
  fs.mkdirSync(path.dirname(patchPath), { recursive: true });
  let patch = `- insert:\n    - id: session-workspace-gate\n      name: '${plugin.replace(/\\/g, '/').replace(/'/g, "''")}'\n`;
  if (localizedNativePicker) {
    patch += `- id: directory-picker\n  disabled: true\n- insert:\n    - id: directory-picker-zh-host\n      name: '${pickerPlugin.replace(/\\/g, '/').replace(/'/g, "''")}'\n    - id: directory-picker-zh-client\n      name: '@deepseek-ai/dsh-client-ui-directory-picker-native'\n`;
  }
  fs.writeFileSync(patchPath, patch, 'utf8');
  dshArgs.push('--patch', patchPath);
}

if (process.env.PORT) {
  dshArgs.push('--port', process.env.PORT);
}
if (process.env.HOST) {
  dshArgs.push('--host', process.env.HOST);
}
if (process.env.DSH_NO_OPEN === '1') {
  dshArgs.push('--no-open');
}
dshArgs.push('--trusted-host', `${host}:${port}`);
const env = {
    ...process.env,
    DEEPSEEK_API_KEY: process.env.DEEPSEEK_API_KEY
  };
if (process.env.WORKSPACE_DIR) {
  env.DSH_HOME = process.env.WORKSPACE_DIR;
}
if (sessionOfficeGate) {
  env.DSH_HOME = path.resolve(process.env.WORKSPACE_DIR || path.join(__dirname, 'workspace'));
  env.DSH_SESSION_OFFICE_GATE = 'enabled';
}
if (isDev) {
  dshArgs.push('--dev');
}

console.log('Starting DeepSeek Harness...');
console.log('Config:', {
  port: process.env.PORT || '3080 (default)',
  host: process.env.HOST || '127.0.0.1 (default)',
  workspace: process.env.WORKSPACE_DIR || './workspace (default)',
  sessionOfficeGate,
  localizedNativePicker,
  apiKey: process.env.DEEPSEEK_API_KEY ? '***SET***' : 'NOT SET (required)'
});

const localCli = path.join(__dirname, 'node_modules', '@deepseek-ai', 'dsh', 'lib', 'bin.js');
if (!fs.existsSync(localCli)) throw new Error('本机 DSH 运行文件不完整；请补齐离线部署包中的 node_modules。');
const dsh = spawn(process.execPath, [localCli, ...dshArgs], {
  stdio: 'inherit',
  shell: false,
  env
});

dsh.on('error', (err) => {
  console.error('Failed to start:', err.message);
  process.exit(1);
});

dsh.on('close', (code) => {
  process.exit(code || 0);
});
