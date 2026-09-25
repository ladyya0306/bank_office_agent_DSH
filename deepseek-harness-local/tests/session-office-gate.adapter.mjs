/** Production-function adapter for session-office-gate.contract.mjs. */
import { spawn } from 'node:child_process';
import { mkdir, writeFile } from 'node:fs/promises';
import { dirname, resolve } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';

const here = dirname(fileURLToPath(import.meta.url));
const production = resolve(here, '../workspace/plugins/session-workspace-gate/lib');
const { handleOnboarding } = await import(pathToFileURL(resolve(production, 'index.js')).href);
const { sameOrChild, stateFile } = await import(pathToFileURL(resolve(production, 'state.js')).href);
const hook = resolve(here, '../../tool/office/hooks/bank_gate.py');
const sessionDirectoriesByHome = new Map();

function runPythonHook(payload, workspaceHome) {
  return new Promise((resolveRun, rejectRun) => {
    const child = spawn('python', [hook], {
      env: { ...process.env, DSH_HOME: workspaceHome, DSH_SESSION_OFFICE_GATE: 'enabled' },
      stdio: ['pipe', 'pipe', 'pipe'],
      windowsHide: true,
    });
    let stdout = '';
    let stderr = '';
    child.stdout.setEncoding('utf8').on('data', (chunk) => { stdout += chunk; });
    child.stderr.setEncoding('utf8').on('data', (chunk) => { stderr += chunk; });
    child.once('error', rejectRun);
    child.once('close', (exitCode) => resolveRun({ exitCode, stdout, stderr }));
    child.stdin.end(JSON.stringify(payload));
  });
}

export async function createSessionOfficeGateTestAdapter({ workspaceHome }) {
  const sessionDirectories = sessionDirectoriesByHome.get(workspaceHome) ?? new Map();
  sessionDirectoriesByHome.set(workspaceHome, sessionDirectories);
  const ctx = {
    sessionController: {
      inspect: async (sessionId) => ({ meta: { cwd: sessionDirectories.get(sessionId) } }),
    },
  };

  async function selectSessionDirectory(sessionId, directory) {
    sessionDirectories.set(sessionId, directory);
  }

  async function selectOfficeWorkspace(sessionId, directory) {
    const response = await handleOnboarding(ctx, new Request('http://test.local/onboarding', {
      method: 'POST',
      headers: { 'content-type': 'application/json' },
      body: JSON.stringify({
        session_id: sessionId,
        session_directory: sessionDirectories.get(sessionId),
        office_workspace: directory,
      }),
    }));
    if (!response.ok) throw new Error((await response.json()).error ?? `HTTP ${response.status}`);
    return response.json();
  }

  return {
    sameOrChild,
    selectSessionDirectory,
    selectOfficeWorkspace,
    runUserPromptHook: (sessionId, prompt) => runPythonHook({
      hook_event_name: 'UserPromptSubmit',
      session_id: sessionId,
      cwd: sessionDirectories.get(sessionId) ?? workspaceHome,
      prompt,
    }, workspaceHome),
    createFreshAdapter: () => createSessionOfficeGateTestAdapter({ workspaceHome }),
    corruptState: (sessionId) => writeFile(stateFile(sessionId), '{not valid json', 'utf8'),
    markLegacy: async (sessionId) => {
      const file = stateFile(sessionId);
      await mkdir(dirname(file), { recursive: true });
      await writeFile(resolve(dirname(file), 'legacy-sessions.json'), JSON.stringify({ version: 1, sessionIds: [sessionId] }), 'utf8');
    },
  };
}
