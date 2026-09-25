import { createHash, randomUUID } from 'node:crypto';

const mappingRevisions = new Map();

function questionVersion(response) {
  return createHash('sha256').update(JSON.stringify({
    task_id: response.task_id,
    status: response.status,
    questions: response.questions || [],
  })).digest('hex');
}

function terminal(response) {
  return ['completed', 'partial', 'needs_mapping', 'failed'].includes(response.status);
}

function validateEnvelope(response) {
  if (!response || typeof response !== 'object' || typeof response.ok !== 'boolean'
      || typeof response.status !== 'string') {
    throw new Error('office.py 返回了无效 JSON 协议响应');
  }
  return response;
}

function preserveTaskContext(response, fallback = {}) {
  const taskId = response.task_id || fallback.task_id;
  const work = response.work || fallback.work;
  const batch = response.batch || fallback.batch;
  const enriched = { ...response };
  if (taskId && !enriched.task_id) enriched.task_id = taskId;
  if (work && !enriched.work) enriched.work = work;
  if (batch && !enriched.batch) enriched.batch = batch;
  if (enriched.status === 'failed' && taskId && typeof enriched.error === 'string'
      && !enriched.error.includes(`任务 ${taskId}`)) {
    enriched.error += `（任务 ${taskId} 可使用 task_id 恢复）`;
  }
  return enriched;
}

function revisionKey(value) {
  return value === undefined || value === null ? null : JSON.stringify(value);
}

function answersFromAsk(response, askResult) {
  const supplied = askResult?.answers;
  if (!Array.isArray(supplied)) return null;
  const asked = new Map((response.questions || []).map((question) => [question.id, question]));
  const answers = [];
  const seen = new Set();
  for (const item of supplied) {
    const question = asked.get(item?.id);
    if (!question) throw new Error('原生提问返回了未知问题，未继续填表');
    if (seen.has(item.id)) throw new Error(`原生提问重复返回问题 ${item.id}，未继续填表`);
    seen.add(item.id);
    const selected = Array.isArray(item.selected) ? item.selected : [];
    if (selected.some((label) => !question.options.some((option) => option.label === label))) {
      throw new Error('原生提问返回了无效选项，未继续填表');
    }
    const custom = typeof item.custom === 'string' ? item.custom.trim() : '';
    if (selected.length > 1 || (selected.length === 1 && custom)) {
      throw new Error(`问题 ${item.id} 必须选择一项或填写自定义答案，不能同时提供多项`);
    }
    answers.push({ id: item.id, selected, custom });
  }
  if (seen.size !== asked.size) return null;
  if (answers.some((item) => item.selected.length !== 1 && !item.custom)) return null;
  return answers;
}

function addTaskId(error, taskId) {
  if (!taskId || error?.task_id === taskId) return error;
  const wrapped = new Error(`${error?.message || String(error)}（任务 ${taskId} 可使用 task_id 恢复）`,
    { cause: error });
  wrapped.task_id = taskId;
  return wrapped;
}

/** Drive only the fixed office.py JSON protocol; all effects are injected for testing. */
export async function runOfficeFill(args, { ask, run, signal }) {
  if (args.mapping_read) {
    if (!args.task_id || args.rule_updates) throw new Error('mapping_read 需要 task_id，且不能与 rule_updates 同传');
    return validateEnvelope(await run({ action: 'read_mapping', work: args.work,
      task_id: args.task_id, mapping_read: args.mapping_read }, { signal }));
  }
  let response;
  try {
    response = preserveTaskContext(validateEnvelope(args.task_id
      ? await run(args.rule_updates
        ? { action: 'update_positions', work: args.work, task_id: args.task_id,
          updates: args.rule_updates }
        : { action: 'status', work: args.work, task_id: args.task_id }, { signal })
      : await run({ action: 'start', work: args.work, source: args.source,
      targets: args.targets, ...(args.batch ? { batch: args.batch } : {}) }, { signal })), {
        task_id: args.task_id, work: args.work, batch: args.batch,
      });
  } catch (error) { throw addTaskId(error, args.task_id); }
  const askedVersions = new Set();
  if (!response.ok && response.status !== 'failed') {
    throw new Error(response.error || '填表任务执行失败');
  }
  if (response.status === 'failed') return response;
  if (args.rule_updates && response.status === 'needs_mapping') {
    const revision = revisionKey(response.mapping_revision);
    if (revision !== null) {
      const key = JSON.stringify([response.work || args.work, response.task_id]);
      const previous = mappingRevisions.get(key);
      if (previous !== undefined && previous === revision) {
        throw new Error(`任务 ${response.task_id} 的位置修复没有推进（mapping_revision 未变化），请先核对 mapping_requests 后再恢复`);
      }
      mappingRevisions.set(key, revision);
      // This is a convenience guard, not an approval store. Keep it bounded.
      if (mappingRevisions.size > 256) mappingRevisions.delete(mappingRevisions.keys().next().value);
    }
  }

  while (!terminal(response)) {
    if (!['awaiting_source', 'awaiting_fill'].includes(response.status)) return response;
    const version = questionVersion(response);
    if (askedVersions.has(version)) {
      throw new Error(`任务 ${response.task_id} 在问题未变化时没有推进；可用 task_id 恢复，系统不会重复弹出同一问题`);
    }
    askedVersions.add(version);
    if (!Array.isArray(response.questions) || response.questions.length === 0) {
      throw new Error(`任务 ${response.task_id} 状态为 ${response.status}，但没有可提问的问题`);
    }
    const questionIds = response.questions.map((question) => question.id);
    if (new Set(questionIds).size !== questionIds.length) {
      throw addTaskId(new Error('待确认问题编号重复，已在弹窗前停止。请修复程序后恢复原任务；重复调用不会解决此错误。'), response.task_id);
    }

    let askResult;
    const waitId = randomUUID();
    const waitStarted = performance.now();
    const waitStartedAt = Date.now() / 1000;
    const waitEvent = (outcome) => ({ id: waitId,
      milliseconds: Math.max(0, performance.now() - waitStarted), outcome,
      started_at: waitStartedAt, ended_at: Date.now() / 1000 });
    const cancelled = async () => {
      let timingRecorded = true;
      try {
        await run({ action: 'record_wait', work: response.work || args.work,
          task_id: response.task_id, user_wait: waitEvent('cancelled') }, { timeoutMs: 5000 });
      } catch { timingRecorded = false; }
      return { ok: true, status: 'cancelled', task_id: response.task_id,
        work: response.work, batch: response.batch,
        ...(!timingRecorded ? { timing_note: '取消等待时长未保存；恢复时该区间标为未细分。' } : {}) };
    };
    try {
      askResult = await ask({ questions: response.questions, signal });
    } catch (error) {
      if (signal?.aborted || error?.name === 'AbortError'
          || error?.message === 'the user cancelled ask_user_question') {
        return cancelled();
      }
      throw addTaskId(error, response.task_id);
    }
    let answers;
    try { answers = answersFromAsk(response, askResult); }
    catch (error) { throw addTaskId(error, response.task_id); }
    if (!answers) return cancelled();

    try {
      const previous = response;
      response = preserveTaskContext(validateEnvelope(await run({ action: 'resume', work: previous.work || args.work,
        task_id: previous.task_id, answers, user_wait: waitEvent('answered') }, { signal })), previous);
    } catch (error) {
      if (signal?.aborted) return { ok: true, status: 'cancelled', task_id: response.task_id,
        work: response.work, batch: response.batch };
      throw addTaskId(error, response.task_id);
    }
    if (!response.ok && response.status !== 'failed') {
      throw new Error(response.error || '填表任务执行失败');
    }
    if (response.status === 'failed') return response;
  }
  return response;
}
