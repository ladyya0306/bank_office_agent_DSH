import { spawn } from 'node:child_process';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { defineTool } from '@deepseek-ai/dsh-tools';
import '@deepseek-ai/dsh-user-questions';
import { runOfficeFill } from './controller.mjs';

export const name = 'office-tool-v2';
export const inject = ['tools', 'userQuestions'];

const pluginDirectory = path.dirname(fileURLToPath(import.meta.url));
const defaultToolRoot = path.resolve(pluginDirectory, '../../../../toolV2');
const MAX_OUTPUT = 2_000_000;

export function runOfficePython(toolRoot, python, payload, { signal, timeoutMs = 10 * 60_000 } = {}) {
  return new Promise((resolve, reject) => {
    if (signal?.aborted) {
      reject(new Error('填表任务已取消；未启动 office.py'));
      return;
    }
    const child = spawn(python, [path.join(toolRoot, 'office.py')], {
      cwd: toolRoot, shell: false, windowsHide: true,
      ...(process.platform === 'win32' ? {} : { detached: true }),
      env: { ...process.env, PYTHONIOENCODING: 'utf-8' },
      stdio: ['pipe', 'pipe', 'pipe'],
    });
    let stdout = '';
    let stderr = '';
    let settled = false;
    let childClosed = false;
    let terminationReason;
    let terminationPromise;
    const timeoutDescription = timeoutMs < 60_000
      ? `${timeoutMs} 毫秒` : `${Math.round(timeoutMs / 60_000)} 分钟`;
    let resolveSpawnAttempt;
    let resolveChildClosed;
    const spawnAttempted = new Promise((resolve) => { resolveSpawnAttempt = resolve; });
    const childClosedPromise = new Promise((resolve) => { resolveChildClosed = resolve; });
    child.once('spawn', resolveSpawnAttempt);
    child.once('error', resolveSpawnAttempt);
    child.once('close', () => { childClosed = true; resolveChildClosed(); });
    const finishError = (error) => {
      if (settled) return;
      settled = true;
      clearTimeout(timer);
      signal?.removeEventListener('abort', abort);
      reject(error);
    };
    async function killProcessTree() {
      if (!child.pid) await spawnAttempted;
      if (childClosed || !child.pid) return;
      if (process.platform === 'win32') {
        await new Promise((resolve, rejectTreeKill) => {
          const killer = spawn('taskkill', ['/PID', String(child.pid), '/T', '/F'], {
            shell: false, windowsHide: true, stdio: ['ignore', 'ignore', 'pipe'],
          });
          let killError = '';
          killer.stderr.setEncoding('utf8');
          killer.stderr.on('data', (chunk) => { killError += chunk; });
          killer.once('error', (error) => rejectTreeKill(new Error(`无法启动 taskkill：${error.message}`)));
          killer.once('close', (code) => {
            if (code === 0 || childClosed) resolve();
            else rejectTreeKill(new Error(`taskkill 未能终止 office.py 进程树（退出码 ${code}）${killError ? `：${killError.trim()}` : ''}`));
          });
        });
      } else {
        try { process.kill(-child.pid, 'SIGKILL'); }
        catch (error) { if (error?.code !== 'ESRCH') throw error; }
      }
      await childClosedPromise;
    }
    function requestTermination(reason) {
      if (settled || terminationPromise) return terminationPromise;
      terminationReason = reason;
      terminationPromise = killProcessTree().then(
        () => finishError(reason),
        (error) => finishError(new Error(`${reason.message}；终止子进程树失败：${error.message}`, { cause: error })),
      );
      return terminationPromise;
    }
    const abort = () => { void requestTermination(new Error('填表任务已取消；可使用 task_id 恢复')); };
    const timer = setTimeout(() => {
      void requestTermination(new Error(`office.py 执行超过 ${timeoutDescription}，已停止进程树`));
    }, timeoutMs);
    signal?.addEventListener('abort', abort, { once: true });
    if (signal?.aborted) abort();
    child.on('error', (error) => finishError(new Error(`无法启动 office.py：${error.message}`)));
    child.stdout.setEncoding('utf8');
    child.stderr.setEncoding('utf8');
    child.stdout.on('data', (chunk) => {
      stdout += chunk;
      if (Buffer.byteLength(stdout, 'utf8') > MAX_OUTPUT) {
        void requestTermination(new Error('office.py 标准输出超过 2 MB 限制，响应已拒绝'));
      }
    });
    child.stderr.on('data', (chunk) => {
      stderr += chunk;
      if (Buffer.byteLength(stderr, 'utf8') > MAX_OUTPUT) {
        void requestTermination(new Error('office.py 错误输出超过 2 MB 限制，响应已拒绝'));
      }
    });
    child.on('close', (code) => {
      if (settled) return;
      if (terminationReason) { finishError(terminationReason); return; }
      settled = true;
      clearTimeout(timer);
      signal?.removeEventListener('abort', abort);
      let envelope;
      try { envelope = JSON.parse(stdout); }
      catch { reject(new Error(`office.py 未返回有效 JSON${stderr ? `：${stderr.trim()}` : ''}`)); return; }
      if (code !== 0 && envelope.ok !== false) {
        reject(new Error(`office.py 退出码为 ${code}${stderr ? `：${stderr.trim()}` : ''}`)); return;
      }
      resolve(envelope);
    });
    child.stdin.on('error', (error) => {
      void requestTermination(new Error(`无法向 office.py 提交任务：${error.message}`));
    });
    child.stdin.end(JSON.stringify(payload));
  });
}

function inside(root, candidate) {
  const relative = path.relative(root, candidate);
  return relative === '' || (relative !== '..' && !relative.startsWith(`..${path.sep}`)
    && !path.isAbsolute(relative));
}

async function resolveWork(work, cwd) {
  if (!cwd) throw new Error('当前会话没有工作目录，无法检查工作区');
  if (typeof work !== 'string' || !work.trim()) throw new Error('请提供工作区 work');
  const fs = await import('node:fs/promises');
  const root = await fs.realpath(cwd);
  const requested = path.resolve(root, work);
  if (!inside(root, requested)) throw new Error('工作区必须位于当前会话目录内');
  const actual = await fs.realpath(requested);
  if (!inside(root, actual)) throw new Error('工作区实际位置越出当前会话目录');
  return actual;
}

export function apply(ctx, config = {}) {
  const toolRoot = path.resolve(config.toolRoot || defaultToolRoot);
  const python = config.python || process.env.DSH_OFFICE_PYTHON
    || (process.platform === 'win32' ? 'python' : 'python3');
  ctx.tools.register(defineTool({
    name: 'office_fill_task',
    description: '用户要求填表时唯一使用的办公工具。用户只需说“填表”；本工具会在必要时原生询问一次并自动继续。恢复中断任务时提供 task_id。终态直接返回文件结果。',
    parameters: {
      work: { type: 'string', required: true, description: '本次任务工作区，必须在当前会话目录内' },
      source: { type: 'array', items: { type: 'string' }, description: '一个或多个信息源文件路径；新建任务必填' },
      targets: { type: 'array', items: { type: 'string' }, description: '要填写的目标模板路径；新建任务必填' },
      batch: { type: 'string', description: '通常省略，程序优先继续相同源文件和目标文件的原任务。仅用户明确指定业务批次时填写 YYYYMMDD-NN，不得自行编造日期或用任务描述代替批次。' },
      task_id: { type: 'string', description: '恢复已有任务时提供；工具会查状态并在需要时继续询问' },
      rule_updates: { type: 'array', items: { type: 'object', additionalProperties: false,
        properties: {
          template: { type: 'string', required: true },
          slot_id: { type: 'string' },
          field: { type: 'string' },
          target: { type: 'object', additionalProperties: true, properties: {} },
          leave_blank: { type: 'boolean' },
          reason: { type: 'string' },
          label: { type: 'string' },
          expected_target: { type: 'object', additionalProperties: true, properties: {} },
        } }, description: '仅对 needs_mapping 的 mapping_requests 提交位置映射：新格式为 template+slot_id+field，或 template+slot_id+leave_blank+reason；兼容旧格式 template+field+target；必须与 task_id 同传' },
    },
    output: { schema: { type: 'object', additionalProperties: true, properties: {} },
      render: (_args, value) => [{ type: 'text', text: JSON.stringify(value) }] },
    async execute(args, exec) {
      if (!args.task_id && (!Array.isArray(args.source) || !Array.isArray(args.targets))) {
        throw new Error('新建填表任务需要 source 和 targets 文件列表');
      }
      if (args.rule_updates && !args.task_id) {
        throw new Error('rule_updates 只能与 task_id 一起用于恢复 needs_mapping 任务');
      }
      const work = await resolveWork(args.work, exec.agent?.session?.header?.cwd);
      return runOfficeFill({ ...args, work }, {
        signal: exec.signal,
        run: (payload, options) => runOfficePython(toolRoot, python, payload, options),
        ask: ({ questions, signal }) => ctx.userQuestions.ask({
          questions, agent: exec.agent, signal,
        }),
      });
    },
  }));
}
