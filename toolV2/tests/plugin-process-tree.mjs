import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import { mkdtemp, mkdir, readFile, rm, writeFile } from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';
import { fileURLToPath, pathToFileURL } from 'node:url';

const toolRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const harnessRoot = process.env.DSH_HARNESS_ROOT || path.resolve(toolRoot, '..', 'deepseek-harness-local');
const pluginParent = path.join(harnessRoot, 'workspace', 'plugins');
const canLoadDshPlugin = await import('node:fs/promises').then(({ access }) =>
  access(path.join(harnessRoot, 'node_modules', '@deepseek-ai', 'dsh-tools')).then(() => true, () => false));
const python = process.env.DSH_OFFICE_PYTHON || (process.platform === 'win32' ? 'python' : 'python3');
const pythonAvailable = spawnSync(python, ['--version'], { windowsHide: true }).status === 0;

function isAlive(pid) {
  try { process.kill(pid, 0); return true; }
  catch (error) { if (error?.code === 'ESRCH') return false; throw error; }
}

async function waitForFile(file, timeoutMs = 5000) {
  const end = Date.now() + timeoutMs;
  while (Date.now() < end) {
    try { return await readFile(file, 'utf8'); }
    catch (error) { if (error?.code !== 'ENOENT') throw error; }
    await new Promise((resolve) => setTimeout(resolve, 25));
  }
  throw new Error('合成 office.py 未能启动子进程');
}

async function waitForExit(pids, timeoutMs = 5000) {
  const end = Date.now() + timeoutMs;
  while (Date.now() < end) {
    if (pids.every((pid) => !isAlive(pid))) return;
    await new Promise((resolve) => setTimeout(resolve, 50));
  }
  const remaining = pids.filter(isAlive);
  assert.deepEqual(remaining, [], `以下合成进程仍在运行：${remaining.join(', ')}`);
}

test('timeout and cancellation stop a synthetic Python parent and child process', {
  skip: !canLoadDshPlugin || !pythonAvailable,
  timeout: 20_000,
}, async () => {
  await mkdir(pluginParent, { recursive: true });
  const pluginCopy = await mkdtemp(path.join(pluginParent, '.office-plugin-test-'));
  const fixtureRoot = await mkdtemp(path.join(os.tmpdir(), 'office-tree-test-'));
  try {
    await writeFile(path.join(pluginCopy, 'index.mjs'),
      await readFile(path.join(toolRoot, 'dsh-plugin', 'index.mjs')));
    await writeFile(path.join(pluginCopy, 'controller.mjs'),
      await readFile(path.join(toolRoot, 'dsh-plugin', 'controller.mjs')));
    await writeFile(path.join(pluginCopy, 'render.mjs'),
      await readFile(path.join(toolRoot, 'dsh-plugin', 'render.mjs')));
    const { runOfficePython } = await import(`${pathToFileURL(path.join(pluginCopy, 'index.mjs')).href}?test=${Date.now()}`);
    const pidFile = path.join(fixtureRoot, 'pids.json');
    const pythonSource = [
      'import json, os, subprocess, sys, time',
      'pid_file = ' + JSON.stringify(pidFile),
      "if len(sys.argv) > 1:",
      '    time.sleep(60)',
      'else:',
      "    flags = getattr(subprocess, 'CREATE_NO_WINDOW', 0)",
      "    child = subprocess.Popen([sys.executable, __file__, 'child'], stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, creationflags=flags)",
      "    with open(pid_file, 'w', encoding='utf-8') as stream: json.dump([os.getpid(), child.pid], stream)",
      '    time.sleep(60)',
    ].join('\n');
    await writeFile(path.join(fixtureRoot, 'office.py'), pythonSource);

    const timeoutRun = runOfficePython(fixtureRoot, python, { action: 'start' }, { timeoutMs: 700 });
    const timeoutPids = JSON.parse(await waitForFile(pidFile));
    await assert.rejects(timeoutRun, /超过 700 毫秒/);
    await waitForExit(timeoutPids);

    await rm(pidFile, { force: true });
    const controller = new AbortController();
    const cancelRun = runOfficePython(fixtureRoot, python, { action: 'start' }, {
      signal: controller.signal, timeoutMs: 10_000,
    });
    const cancelPids = JSON.parse(await waitForFile(pidFile));
    controller.abort();
    await assert.rejects(cancelRun, /任务已取消/);
    await waitForExit(cancelPids);
  } finally {
    await rm(pluginCopy, { recursive: true, force: true });
    await rm(fixtureRoot, { recursive: true, force: true });
  }
});
