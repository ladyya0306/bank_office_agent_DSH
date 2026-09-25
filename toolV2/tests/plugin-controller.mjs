import assert from 'node:assert/strict';
import test from 'node:test';
import { runOfficeFill } from '../dsh-plugin/controller.mjs';

const pending = (status = 'awaiting_source', question = 'q1') => ({
  ok: true, status, task_id: 'task-1', work: 'work', batch: 'b1',
  questions: [{ id: question, header: '源文件', question: '请选择源文件',
    options: [{ label: 'source.xlsx', description: '候选文件' }] }],
});
const completed = () => ({ ok: true, status: 'completed', task_id: 'task-1', work: 'work',
  batch: 'b1', questions: [], results: [{ file: 'out.xlsx' }] });
const answer = (id = 'q1') => ({ answers: [{ id, selected: ['source.xlsx'], custom: '' }] });

test('duplicate question identities stop before displaying a broken popup', async () => {
  let asked = 0;
  let runs = 0;
  const response = pending();
  response.questions.push({ ...response.questions[0] });
  await assert.rejects(() => runOfficeFill({ work: 'work', task_id: 'task-1' }, {
    ask: async () => { asked += 1; return answer(); },
    run: async () => { runs += 1; return response; },
  }), (error) => error.message.includes('弹窗前停止') && error.task_id === 'task-1');
  assert.equal(asked, 0);
  assert.equal(runs, 1);
});

test('asks once and automatically resumes through source and fill questions', async () => {
  const calls = [];
  const asks = [];
  const responses = [pending(), pending('awaiting_fill', 'q2'), completed()];
  const result = await runOfficeFill({ work: 'work', source: ['in.xlsx'], targets: ['out.xlsx'] }, {
    ask: async ({ questions }) => { asks.push(questions[0].id); return answer(questions[0].id); },
    run: async (payload) => { calls.push(payload); return responses.shift(); },
  });
  assert.equal(result.status, 'completed');
  assert.deepEqual(asks, ['q1', 'q2']);
  assert.deepEqual(calls.map((call) => call.action), ['start', 'resume', 'resume']);
  assert.equal(calls[1].answers[0].id, 'q1');
});

test('cancel returns only the resumable task identity, never synthetic answers', async () => {
  const calls = [];
  const result = await runOfficeFill({ work: 'work', source: ['in.xlsx'], targets: ['out.xlsx'] }, {
    ask: async () => ({ cancelled: true }),
    run: async (payload) => { calls.push(payload); return pending(); },
  });
  assert.deepEqual(result, { ok: true, status: 'cancelled', task_id: 'task-1',
    work: 'work', batch: 'b1' });
  assert.equal(calls.length, 2);
  assert.equal(calls[1].action, 'record_wait');
  assert.equal(calls[1].user_wait.outcome, 'cancelled');
  assert.equal(calls[1].answers, undefined);
});

test('recognizes DSH cancellation error without swallowing unrelated ask errors', async () => {
  const result = await runOfficeFill({ work: 'work', source: ['in.xlsx'], targets: ['out.xlsx'] }, {
    ask: async () => { throw new Error('the user cancelled ask_user_question'); },
    run: async () => pending(),
  });
  assert.equal(result.status, 'cancelled');
  assert.equal(result.task_id, 'task-1');
  await assert.rejects(() => runOfficeFill({ work: 'work', source: ['in.xlsx'], targets: ['out.xlsx'] }, {
    ask: async () => { throw new Error('question service unavailable'); },
    run: async () => pending(),
  }), (error) => error.message.includes('question service unavailable'));
  await assert.rejects(() => runOfficeFill({ work: 'work', task_id: 'task-ask-error' }, {
    ask: async () => { throw new Error('question service unavailable'); },
    run: async () => ({ ...pending(), task_id: 'task-ask-error' }),
  }), (error) => error.message.includes('question service unavailable')
      && error.message.includes('task-ask-error'));
});

test('does not resume on duplicate, missing, or unanswered question replies', async () => {
  for (const askResult of [
    { answers: [] },
    { answers: [{ id: 'q1', selected: [] }] },
  ]) {
    const calls = [];
    const result = await runOfficeFill({ work: 'work', source: ['in.xlsx'], targets: ['out.xlsx'] }, {
      ask: async () => askResult,
      run: async (payload) => { calls.push(payload); return pending(); },
    });
    assert.equal(result.status, 'cancelled');
    assert.deepEqual(calls.map(x => x.action), ['start', 'record_wait'], 'incomplete answer must not be sent to resume');
  }
  await assert.rejects(() => runOfficeFill({ work: 'work', source: ['in.xlsx'], targets: ['out.xlsx'] }, {
    ask: async () => ({ answers: [
      { id: 'q1', selected: ['source.xlsx'] }, { id: 'q1', selected: ['source.xlsx'] },
    ] }),
    run: async () => pending(),
  }), /重复返回问题/);
});

test('mapping pages use the existing task without asking or advancing', async () => {
  const calls = [];
  const result = await runOfficeFill({ work: 'work', task_id: 'task-1',
    mapping_read: { section: 'positions', template: 'one.xlsx', offset: 5 } }, {
    ask: async () => { throw new Error('must not ask'); },
    run: async payload => { calls.push(payload); return { ok: true, status: 'needs_mapping', mapping_page: { items: [] } }; },
  });
  assert.equal(result.status, 'needs_mapping');
  assert.deepEqual(calls, [{ action: 'read_mapping', work: 'work', task_id: 'task-1',
    mapping_read: { section: 'positions', template: 'one.xlsx', offset: 5 } }]);
});

test('native waiting is telemetry only and never a model supplied answer', async () => {
  const calls = [];
  await runOfficeFill({ work: 'work', task_id: 'task-1' }, {
    ask: async () => answer(),
    run: async payload => { calls.push(payload); return payload.action === 'status' ? pending() : completed(); },
  });
  assert.equal(calls[1].user_wait.outcome, 'answered');
  assert.ok(calls[1].user_wait.id);
  assert.ok(calls[1].user_wait.milliseconds >= 0);
});

test('rejects multi-select answers for single choice and sends the supported one-source array', async () => {
  await assert.rejects(() => runOfficeFill({ work: 'work', source: ['in.xlsx'], targets: ['out.xlsx'] }, {
    ask: async () => ({ answers: [{ id: 'q1', selected: ['source.xlsx', 'source.xlsx'] }] }),
    run: async () => pending(),
  }), /必须选择一项/);
  let startPayload;
  await runOfficeFill({ work: 'work', source: ['in.xlsx'], targets: ['out.xlsx'] }, {
    ask: async () => answer(),
    run: async (payload) => {
      if (payload.action === 'start') { startPayload = payload; return completed(); }
      return completed();
    },
  });
  assert.deepEqual(startPayload.source, ['in.xlsx']);
});

test('submits constrained mapping updates only with task_id, then continues natively', async () => {
  const calls = [];
  const result = await runOfficeFill({ work: 'work', task_id: 'task-1', rule_updates: [{
    template: 'out.xlsx', field: '日期', target: { sheet: 'Sheet1', cell: 'B2' },
  }] }, {
    ask: async ({ questions }) => answer(questions[0].id),
    run: async (payload) => {
      calls.push(payload);
      if (payload.action === 'update_positions') return pending('awaiting_fill');
      return completed();
    },
  });
  assert.equal(result.status, 'completed');
  assert.equal(calls[0].action, 'update_positions');
  assert.deepEqual(calls[0].updates[0].target, { sheet: 'Sheet1', cell: 'B2' });
  assert.equal(calls[1].action, 'resume');
});

test('allows a first mapping response that still needs mapping', async () => {
  const result = await runOfficeFill({ work: 'work', task_id: 'task-mapping', rule_updates: [{
    template: 'out.xlsx', field: '日期', target: { sheet: 'Sheet1', cell: 'B2' },
  }] }, {
    ask: async () => answer(), run: async () => ({ ok: true, status: 'needs_mapping',
      task_id: 'task-mapping', work: 'work', batch: 'b1', questions: [], results: [],
      mapping_revision: 'r1' }),
  });
  assert.equal(result.status, 'needs_mapping');
  await assert.rejects(() => runOfficeFill({ work: 'work', task_id: 'task-mapping', rule_updates: [{
    template: 'out.xlsx', field: '日期', target: { sheet: 'Sheet1', cell: 'B2' },
  }] }, {
    ask: async () => answer(), run: async () => ({ ok: true, status: 'needs_mapping',
      task_id: 'task-mapping', work: 'work', batch: 'b1', questions: [], results: [],
      mapping_revision: 'r1' }),
  }), /mapping_revision 未变化/);
});

test('recovery that returns the same unanswered question fails without asking twice', async () => {
  let askCount = 0;
  let runCount = 0;
  await assert.rejects(() => runOfficeFill({ work: 'work', task_id: 'task-1' }, {
    ask: async () => { askCount += 1; return answer(); },
    run: async (payload) => {
      runCount += 1;
      return payload.action === 'status' ? pending() : pending();
    },
  }), /问题未变化时没有推进/);
  assert.equal(askCount, 1);
  assert.equal(runCount, 2);
});

test('propagates execution errors and can resume an existing task', async () => {
  await assert.rejects(() => runOfficeFill({ work: 'work', source: [], targets: [] }, {
    ask: async () => answer(), run: async () => { throw new Error('python failed'); },
  }), /python failed/);
  await assert.rejects(() => runOfficeFill({ work: 'work', task_id: 'task-resume' }, {
    ask: async () => answer(), run: async () => { throw new Error('python failed'); },
  }), (error) => error.message.includes('python failed')
      && error.message.includes('task-resume') && error.task_id === 'task-resume');
  const calls = [];
  const result = await runOfficeFill({ work: 'work', task_id: 'task-1' }, {
    ask: async ({ questions }) => answer(questions[0].id),
    run: async (payload) => { calls.push(payload); return calls.length === 1 ? pending() : completed(); },
  });
  assert.equal(result.status, 'completed');
  assert.equal(calls[0].action, 'status');
  assert.equal(calls[1].action, 'resume');
});

test('preserves task context when resume returns a failed envelope without identity', async () => {
  let calls = 0;
  const result = await runOfficeFill({ work: 'work', source: ['in.xlsx'], targets: ['out.xlsx'] }, {
    ask: async () => answer(),
    run: async (payload) => {
      calls += 1;
      return calls === 1 ? pending() : { ok: false, status: 'failed', questions: [], results: [],
        error: '回答保存失败' };
    },
  });
  assert.equal(result.status, 'failed');
  assert.equal(result.task_id, 'task-1');
  assert.equal(result.work, 'work');
  assert.equal(result.batch, 'b1');
  assert.match(result.error, /回答保存失败/);
  assert.match(result.error, /task_id 恢复/);
  assert.equal(calls, 2);
});

test('preserves an existing task id when the initial status call fails', async () => {
  const result = await runOfficeFill({ work: 'work', task_id: 'task-resume' }, {
    ask: async () => answer(),
    run: async () => ({ ok: false, status: 'failed', questions: [], results: [], error: '状态读取失败' }),
  });
  assert.equal(result.status, 'failed');
  assert.equal(result.task_id, 'task-resume');
  assert.match(result.error, /状态读取失败/);
  assert.match(result.error, /task_id 恢复/);
});
