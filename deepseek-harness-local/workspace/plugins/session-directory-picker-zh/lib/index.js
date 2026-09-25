// Local Windows title adapter for @deepseek-ai/dsh-host-directory-picker-native
// (MIT). The official /worker entry continues to own the COM dialog itself.
import { DirectoryPicker } from '@deepseek-ai/dsh-host-directory-picker';
import { pickNativeDirectory } from '@deepseek-ai/dsh-host-directory-picker-native';
import { spawn } from 'node:child_process';
import { fileURLToPath } from 'node:url';

export const DIALOG_TITLE = '选择本地文件夹';
const WM_CLOSE = 16;
const CLOSE_RETRY_MS = 150;
const CLOSE_MAX_ATTEMPTS = 20;

function spawnDialogWorker({ title }) {
  return spawn(process.execPath, [fileURLToPath(import.meta.resolve('@deepseek-ai/dsh-host-directory-picker-native/worker'))], {
    env: { ...process.env, DSH_DIALOG_TITLE: title },
    stdio: ['ignore', 'inherit', 'inherit', 'ipc'],
    windowsHide: true,
  });
}

async function closeThreadWindows(threadId) {
  const koffi = (await import('koffi')).default;
  const user32 = koffi.load('user32.dll');
  const enumThreadWindows = user32.func('__stdcall', 'EnumThreadWindows', 'int', ['uint32', 'void *', 'intptr']);
  const postMessageW = user32.func('__stdcall', 'PostMessageW', 'int', ['void *', 'uint32', 'uintptr', 'intptr']);
  const callbackType = koffi.proto('int __stdcall DshEnumThreadWndProc(void *hwnd, intptr lparam)');
  const callback = koffi.register((window) => {
    postMessageW(window, WM_CLOSE, 0, 0);
    return 1;
  }, koffi.pointer(callbackType));
  try {
    enumThreadWindows(threadId, callback, 0);
  } finally {
    koffi.unregister(callback);
  }
}

export async function pickLocalizedWin32Directory(signal, internals = {}) {
  if (signal.aborted) throw new Error('native directory picker aborted');
  const worker = (internals.spawnWorker ?? spawnDialogWorker)({ title: DIALOG_TITLE });
  const closeWindows = internals.closeThreadWindows ?? closeThreadWindows;
  const closeRetryMs = internals.closeRetryMs ?? CLOSE_RETRY_MS;
  let dialogThreadId;
  let closeTimer;
  let settled = false;
  return new Promise((resolve, reject) => {
    const settle = (action) => {
      if (settled) return;
      settled = true;
      if (closeTimer !== undefined) clearInterval(closeTimer);
      signal.removeEventListener('abort', onAbort);
      worker.unref?.();
      action();
    };
    const postClose = () => {
      if (dialogThreadId !== undefined) closeWindows(dialogThreadId).catch(() => {});
    };
    const onAbort = () => {
      let attempts = 0;
      closeTimer = setInterval(() => {
        attempts += 1;
        if (attempts > CLOSE_MAX_ATTEMPTS) {
          settle(() => {
            worker.kill();
            reject(new Error('native directory picker aborted (dialog unresponsive; worker killed)'));
          });
          return;
        }
        postClose();
      }, closeRetryMs);
      postClose();
    };
    signal.addEventListener('abort', onAbort, { once: true });
    if (signal.aborted) onAbort();
    worker.on('message', (message) => {
      switch (message?.kind) {
        case 'showing':
          dialogThreadId = message.threadId;
          if (signal.aborted) postClose();
          return;
        case 'done':
          settle(() => signal.aborted ? reject(new Error('native directory picker aborted')) : resolve(message.path));
          return;
        case 'error':
          settle(() => reject(new Error(`win32 folder dialog failed: ${message.message}`)));
          return;
        default:
          settle(() => reject(new TypeError(`unknown win32 dialog worker message kind: ${String(message?.kind)}`)));
      }
    });
    worker.on('error', (error) => settle(() => reject(error)));
    worker.on('exit', () => settle(() => reject(new Error('win32 folder dialog worker exited before reporting a result'))));
  });
}

export default class LocalizedNativeDirectoryPicker extends DirectoryPicker {
  nativeCapability = {
    kind: 'native',
    pick: (signal) => process.platform === 'win32'
      ? pickLocalizedWin32Directory(signal)
      : pickNativeDirectory(signal),
  };

  capability() {
    return this.nativeCapability;
  }
}
