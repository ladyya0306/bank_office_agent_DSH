/**
 * Contract for the model-visible SessionStart instruction after web onboarding.
 * This cannot prove a model's future wording; it prevents the hook itself from
 * reintroducing instructions that cause a second directory-confirmation ask.
 */
import assert from 'node:assert/strict';
import { spawn } from 'node:child_process';
import { readFile } from 'node:fs/promises';

const hook = 'D:\\DSH\\tool\\office\\hooks\\bank_gate.py';
const client = new URL('../workspace/plugins/session-workspace-gate/lib/client.js', import.meta.url);

function runHook(payload) {
  return new Promise((resolve, reject) => {
    const process = spawn('python', [hook], { stdio: ['pipe', 'pipe', 'pipe'], windowsHide: true });
    let stdout = '';
    let stderr = '';
    process.stdout.setEncoding('utf8').on('data', (chunk) => { stdout += chunk; });
    process.stderr.setEncoding('utf8').on('data', (chunk) => { stderr += chunk; });
    process.once('error', reject);
    process.once('close', (exitCode) => resolve({ exitCode, stdout, stderr }));
    process.stdin.end(JSON.stringify(payload));
  });
}

const result = await runHook({ hook_event_name: 'SessionStart', session_id: 'discipline-session-01' });
assert.equal(result.exitCode, 0, `SessionStart hook 应正常返回：${result.stderr}`);
const output = JSON.parse(result.stdout);
const context = output?.hookSpecificOutput?.additionalContext;
assert.equal(output?.hookSpecificOutput?.hookEventName, 'SessionStart');
assert.equal(typeof context, 'string', 'SessionStart 必须输出模型可见的纪律文本');
assert.match(context, /网页.*确认/, '纪律文本必须说明网页已完成目录确认');
assert.match(context, /无需再次询问/, '纪律文本必须禁止重复确认提问');
assert.match(context, /ask_user_question/, '纪律文本必须点名禁止用 ask_user_question 重问');
assert.doesNotMatch(context, /新会话须先选会话目录/, '不得保留已完成后仍要求选择目录的旧指令');
assert.doesNotMatch(context, /办公工作区须在会话开始前确认/, '不得保留已完成后仍要求确认工作区的旧指令');

const clientSource = await readFile(client, 'utf8');
assert.doesNotMatch(clientSource, /ask_user_question/, '新 UI 不得自行调用 ask_user_question 进行目录确认');
console.log('RESULT: SessionStart 纪律文本与客户端均未引入重复目录确认调用');
