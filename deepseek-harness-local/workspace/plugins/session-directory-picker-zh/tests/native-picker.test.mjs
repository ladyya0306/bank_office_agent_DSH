import assert from 'node:assert/strict';
import { EventEmitter } from 'node:events';
import test from 'node:test';
import { pickLocalizedWin32Directory } from '../lib/index.js';

function fakeWorker() {
  const worker = new EventEmitter();
  worker.unref = () => {};
  worker.kill = () => { worker.killed = true; };
  return worker;
}

test('passes the Chinese title to the worker and returns the chosen directory', async () => {
  const worker = fakeWorker();
  let actualTitle;
  const picked = pickLocalizedWin32Directory(new AbortController().signal, {
    spawnWorker: ({ title }) => { actualTitle = title; return worker; },
  });
  assert.equal(actualTitle, '选择本地文件夹');
  worker.emit('message', { kind: 'done', path: 'D:\\Work' });
  assert.equal(await picked, 'D:\\Work');
});

test('returns null when the operator cancels the dialog', async () => {
  const worker = fakeWorker();
  const picked = pickLocalizedWin32Directory(new AbortController().signal, {
    spawnWorker: () => worker,
  });
  worker.emit('message', { kind: 'done', path: null });
  assert.equal(await picked, null);
});

test('closes the native dialog thread on abort and rejects the pick', async () => {
  const controller = new AbortController();
  const worker = fakeWorker();
  const closedThreads = [];
  const picked = pickLocalizedWin32Directory(controller.signal, {
    spawnWorker: () => worker,
    closeThreadWindows: async (threadId) => { closedThreads.push(threadId); },
    closeRetryMs: 1000,
  });
  worker.emit('message', { kind: 'showing', threadId: 42 });
  controller.abort();
  worker.emit('message', { kind: 'done', path: null });
  await assert.rejects(picked, /aborted/);
  assert.deepEqual(closedThreads, [42]);
});

test('reports worker exit without a result', async () => {
  const worker = fakeWorker();
  const picked = pickLocalizedWin32Directory(new AbortController().signal, {
    spawnWorker: () => worker,
  });
  worker.emit('exit', 1);
  await assert.rejects(picked, /exited before reporting/);
});
