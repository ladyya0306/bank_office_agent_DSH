import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import { mkdtemp, mkdir, access, rm } from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const harness = process.env.DSH_HARNESS_ROOT || path.resolve(root, '../deepseek-harness-local');
const { apply } = await import(pathToFileURL(path.join(harness, 'workspace/plugins/office-tool-v2/index.mjs')).href);
const python = process.env.DSH_OFFICE_PYTHON || (process.platform === 'win32' ? 'python' : 'python3');
const temporary = await mkdtemp(path.join(os.tmpdir(), 'toolv2-workspaces-'));
let tool;
let questionsAsked = 0;
apply({ tools: { register: (definition) => { tool = definition; } },
  userQuestions: { ask: async ({ questions }) => {
    questionsAsked += questions.length;
    return { answers: questions.map(q => ({ id: q.id, selected: [q.options[0].label], custom: '' })) };
  } } }, { toolRoot: root, python });

function pythonRun(script, args) {
  const run = spawnSync(python, ['-', ...args], { input: script, encoding: 'utf8',
    windowsHide: true, env: { ...process.env, PYTHONUTF8: '1' } });
  assert.equal(run.status, 0, run.stderr);
  return run.stdout;
}

try {
  for (let i = 0; i < 2; i++) {
    const cwd = path.join(temporary, `随机项目 ${i + 1}`);
    const work = i === 0 ? cwd : path.join(cwd, '另选的办公区');
    await mkdir(work, { recursive: true });
    const value = `98765432${i}`;
    pythonRun('import sys\nfrom pathlib import Path\nfrom docx import Document\np=Path(sys.argv[1])\n(p/"原始资料").mkdir()\n(p/"待填写表格").mkdir()\nd=Document()\nd.add_paragraph("借款人：合成企业")\nd.add_paragraph("收款账号："+sys.argv[2])\nd.save(p/"原始资料"/"材料.docx")\nd=Document()\nd.add_paragraph("收款账号：")\nd.save(p/"待填写表格"/"目标.docx")', [work, value]);
    const args = { work, source: [path.join(work, '原始资料', '材料.docx')],
      targets: [path.join(work, '待填写表格', '目标.docx')], batch: '20260925-01' };
    const before = questionsAsked;
    const first = await tool.execute(args, { agent: { session: { header: { cwd, id: `new-${i}` } } } });
    assert.equal(first.status, 'completed', JSON.stringify(first));
    assert.ok(questionsAsked > before, 'A different workspace must not inherit customer answers');
    await access(path.join(work, 'db', 'workflow.db'));
    const output = first.results[0].output;
    assert.ok(!path.relative(work, output).startsWith('..'));
    assert.ok(pythonRun('import sys\nfrom docx import Document\nprint("\\n".join(p.text for p in Document(sys.argv[1]).paragraphs))', [output]).includes(value));
    const answered = questionsAsked;
    const second = await tool.execute(args, { agent: { session: { header: { cwd, id: `return-${i}` } } } });
    assert.equal(second.status, 'completed');
    assert.equal(questionsAsked, answered);
    assert.equal(second.counters.fill_processes, first.counters.fill_processes);
  }
  console.log(JSON.stringify({ passed: true, randomWorkspacesOutsideInstallation: 2,
    chineseAndSpacePaths: true, separateOfficeSubfolder: true, newSessionReusesAnswers: true,
    independentDatabases: true, realWordContentsVerified: true, nativeAnswersSimulated: true }));
} finally {
  assert.equal(path.dirname(temporary), os.tmpdir());
  assert.ok(path.basename(temporary).startsWith('toolv2-workspaces-'));
  await rm(temporary, { recursive: true, force: true });
}
