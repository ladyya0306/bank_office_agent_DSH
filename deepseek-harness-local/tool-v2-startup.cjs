// Portable local startup: no downloads, no business-workspace mutations.
const fs = require('node:fs');
const path = require('node:path');
const { spawnSync } = require('node:child_process');

function prepareToolV2(harnessRoot, env = process.env) {
  const toolRoot = path.resolve(harnessRoot, '../toolV2');
  const installer = path.join(toolRoot, 'install.py');
  if (!fs.existsSync(installer)) {
    throw new Error(`缺少 toolV2，请将整个 toolV2 文件夹与 deepseek-harness-local 放在同一父目录：${toolRoot}`);
  }
  const python = env.DSH_OFFICE_PYTHON || (process.platform === 'win32' ? 'python' : 'python3');
  const result = spawnSync(python, [installer, '--dsh-root', harnessRoot], {
    cwd: harnessRoot, env: { ...env, PYTHONIOENCODING: 'utf-8', PYTHONUTF8: '1' },
    encoding: 'utf8', shell: false, windowsHide: true, timeout: 30000,
  });
  if (result.error) throw new Error(`toolV2 启动检查失败：${result.error.message}`);
  let reply;
  try { reply = JSON.parse(result.stdout); }
  catch { throw new Error(`toolV2 启动检查没有返回有效结果：${result.stderr || result.stdout}`); }
  if (result.status !== 0 || !reply.ok) throw new Error(reply.error || 'toolV2 接入失败');
  return reply;
}

function portableWorkspace(harnessRoot, configured) {
  const standard = path.join(harnessRoot, 'workspace');
  if (!configured || /(?:^|\/)deepseek-harness-local\/workspace\/?$/i.test(configured.replace(/\\/g, '/'))) {
    return standard;
  }
  return path.resolve(harnessRoot, configured);
}

module.exports = { prepareToolV2, portableWorkspace };
