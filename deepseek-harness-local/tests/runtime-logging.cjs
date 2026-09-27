const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { spawnSync } = require('node:child_process');
const { openRuntimeLogs } = require('../runtime-logging.cjs');

function fixture(t) {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), 'dsh-log-test-'));
  t.after(() => {
    const resolved = fs.realpathSync(root);
    assert.equal(path.dirname(resolved).toLowerCase(), fs.realpathSync(os.tmpdir()).toLowerCase());
    assert.ok(path.basename(resolved).startsWith('dsh-log-test-'));
    fs.rmSync(resolved, { recursive: true });
  });
  return root;
}

test('preserves UTF-8 byte chunks and stderr, mirrors output, never overwrites a previous run', t => {
  const root = fixture(t), out = [], err = [];
  const mirrors = { stdout: { write: b => out.push(b) }, stderr: { write: b => err.push(b) } };
  const first = openRuntimeLogs(root, mirrors);
  const bytes = Buffer.from('合成日志：800万元\n');
  first.stdout(bytes.subarray(0, 2)); first.stdout(bytes.subarray(2));
  first.error('fixture-error'); first.close(); first.close();
  const second = openRuntimeLogs(root, mirrors); second.info('second'); second.close();
  assert.notEqual(first.directory, second.directory);
  assert.equal(fs.readFileSync(path.join(first.directory, 'stdout.log'), 'utf8'), bytes.toString());
  assert.equal(Buffer.concat(out).toString(), bytes.toString() + 'second\n');
  assert.equal(Buffer.concat(err).toString(), 'fixture-error\n');
  assert.equal(fs.readFileSync(path.join(first.directory, 'stderr.log'), 'utf8'), 'fixture-error\n');
  assert.throws(() => first.stdout('late'), /already closed/);
  assert.deepEqual(fs.readdirSync(root), ['logs']);
});

function installFakeLauncher(root) {
  const harness = path.join(root, 'harness'); fs.mkdirSync(harness);
  for (const name of ['start.js', 'runtime-logging.cjs']) {
    fs.copyFileSync(path.join(__dirname, '..', name), path.join(harness, name));
  }
  fs.writeFileSync(path.join(harness, 'tool-v2-startup.cjs'), `exports.prepareToolV2=()=>({changed:false}); exports.portableWorkspace=root=>root;`);
  const dotenv = path.join(harness, 'node_modules', 'dotenv'); fs.mkdirSync(dotenv, { recursive: true });
  fs.writeFileSync(path.join(dotenv, 'index.js'), 'exports.config=()=>({});');
  const cli = path.join(harness, 'node_modules', '@deepseek-ai', 'dsh', 'lib'); fs.mkdirSync(cli, { recursive: true });
  fs.writeFileSync(path.join(cli, 'bin.js'), `process.stdout.write('fake-child-out\\n'); process.stderr.write('fake-child-error\\n'); process.exitCode=7;`);
  return harness;
}

test('real launcher captures child streams and exit failure without depending on caller cwd', t => {
  const root = fixture(t), harness = installFakeLauncher(root);
  const result = spawnSync(process.execPath, [path.join(harness, 'start.js')], {
    cwd: root, encoding: 'utf8', timeout: 10000,
    env: { SystemRoot: process.env.SystemRoot || '', PATH: process.env.PATH || '' },
  });
  assert.equal(result.error, undefined); assert.equal(result.status, 7);
  const base = path.join(harness, 'logs', 'runtime');
  const dirs = fs.readdirSync(base); assert.equal(dirs.length, 1);
  const stdout = fs.readFileSync(path.join(base, dirs[0], 'stdout.log'), 'utf8');
  assert.match(stdout, /toolV2 接入检查通过/); assert.match(stdout, /fake-child-out/);
  assert.match(result.stdout, /fake-child-out/); assert.match(result.stderr, /fake-child-error/);
  assert.match(fs.readFileSync(path.join(base, dirs[0], 'stderr.log'), 'utf8'), /fake-child-error/);
  assert.equal(fs.existsSync(path.join(root, 'logs')), false);
});

test('an invalid log directory fails visibly before launching the child', t => {
  const root = fixture(t), harness = installFakeLauncher(root);
  fs.writeFileSync(path.join(harness, 'logs'), 'existing file');
  const result = spawnSync(process.execPath, [path.join(harness, 'start.js')], {
    cwd: root, encoding: 'utf8', timeout: 10000,
    env: { SystemRoot: process.env.SystemRoot || '', PATH: process.env.PATH || '' },
  });
  assert.notEqual(result.status, 0); assert.match(result.stderr, /ENOTDIR|EEXIST/);
  assert.doesNotMatch(result.stdout, /fake-child-out/);
  assert.equal(fs.readFileSync(path.join(harness, 'logs'), 'utf8'), 'existing file');
});
