import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import { mkdtemp, rm } from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { runOfficeFill } from '../dsh-plugin/controller.mjs';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const harness = process.env.DSH_HARNESS_ROOT || path.resolve(root, '../deepseek-harness-local');
const plugin = path.join(harness, 'workspace/plugins/office-tool-v2/index.mjs');
const { runOfficePython } = await import(pathToFileURL(plugin).href);
const work = await mkdtemp(path.join(os.tmpdir(), 'office-v2-integration-'));
const python = process.env.DSH_OFFICE_PYTHON || (process.platform === 'win32' ? 'python' : 'python3');
try {
  const setup = spawnSync(python, ['-', work], {encoding: 'utf8', windowsHide: true, input: [
    'import sys', 'from pathlib import Path', 'from docx import Document',
    'p=Path(sys.argv[1])', '(p/"source").mkdir()', '(p/"target").mkdir()', 'd=Document()',
    'd.add_paragraph("收款账号：123456789")', 'd.add_paragraph("借款人：合成企业")',
    'd.save(p/"source"/"source.docx")', 'd=Document()', 'd.add_paragraph("收款账号：")',
    'd.save(p/"target"/"target.docx")',
  ].join('\n'), env: {...process.env, PYTHONUTF8: '1'}});
  assert.equal(setup.status, 0, setup.stderr);
  const args = {work, source: [path.join(work, 'source', 'source.docx')], targets: [path.join(work, 'target', 'target.docx')], batch: '20260925-01'};
  let cards = 0;
  const ports = {
    run: (payload, options) => runOfficePython(root, python, payload, options),
    ask: async ({questions}) => {
      cards += questions.length;
      return {answers: questions.map(q => ({id:q.id, selected:[q.options.find(o=>o.label.startsWith('归属：合成企业'))?.label || q.options[0].label], custom:''}))};
    },
  };
  const cancelled = await runOfficeFill(args, { ...ports, ask: async () => {
    await new Promise(resolve => setTimeout(resolve, 20));
    return { cancelled: true };
  } });
  assert.equal(cancelled.status, 'cancelled');
  const first = await runOfficeFill(args, ports);
  assert.equal(first.status, 'completed', JSON.stringify(first));
  assert.equal(first.results.length, 1);
  assert.ok(cards > 0);
  assert.equal(first.timing.cancelled_waits, 1);
  assert.equal(first.timing.stages.user_confirmation_wait.count, 2);
  assert.ok(first.timing.stages.user_confirmation_wait.seconds >= 0.02);
  const originalCards = cards;
  const second = await runOfficeFill({work, task_id: first.task_id}, ports);
  assert.equal(second.status, 'completed');
  assert.equal(cards, originalCards, 'Unchanged task asked again');
  assert.equal(second.counters.fill_processes, first.counters.fill_processes);
  assert.deepEqual(second.timing.stages.user_confirmation_wait, first.timing.stages.user_confirmation_wait);
  const changed = spawnSync(python, ['-', work], {encoding:'utf8', windowsHide:true,
    env:{...process.env, PYTHONUTF8:'1'}, input:[
      'import sys', 'from pathlib import Path', 'from docx import Document',
      'd=Document()', 'd.add_paragraph("收款账号：987654321")',
      'd.add_paragraph("借款人：合成企业")', 'd.save(Path(sys.argv[1])/"source"/"source.docx")',
    ].join('\n')});
  assert.equal(changed.status, 0, changed.stderr);
  const third = await runOfficeFill(args, ports);
  assert.equal(third.status, 'completed', JSON.stringify(third));
  assert.equal(third.task_id, first.task_id);
  assert.equal(third.counters.fill_processes, first.counters.fill_processes + 1);
  assert.notEqual(third.results[0].output, first.results[0].output);
  const readBack = spawnSync(python, ['-', third.results[0].output], {encoding:'utf8', windowsHide:true,
    input:'import sys\nfrom docx import Document\nprint("\\n".join(p.text for p in Document(sys.argv[1]).paragraphs))',
    env:{...process.env, PYTHONUTF8:'1'}});
  assert.equal(readBack.status, 0, readBack.stderr);
  assert.ok(readBack.stdout.includes('987654321'));
  console.log(JSON.stringify({passed: true, nativeAskInterfaceSimulated: true,
    pythonWorkerReal: true, firstQuestions: originalCards, resumeQuestions: 0,
    changedSourceQuestions: cards-originalCards, sameBatchNewValueVerified:true,
    cancelledWaitRecorded: true, resumedWaitNotDoubleCounted: true}));
} finally {
  assert.ok(path.dirname(work) === os.tmpdir() && path.basename(work).startsWith('office-v2-integration-'));
  await rm(work, {recursive: true, force: true});
}
