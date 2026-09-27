// Launcher logs only; DSH session and business audit records stay separate.
const fs = require('node:fs');
const path = require('node:path');
const { format } = require('node:util');

function openRuntimeLogs(harnessRoot, mirrors = { stdout: process.stdout, stderr: process.stderr }) {
  const base = path.join(harnessRoot, 'logs', 'runtime');
  fs.mkdirSync(base, { recursive: true });
  const stamp = new Date().toISOString().replace(/[:.]/g, '-');
  const directory = fs.mkdtempSync(path.join(base, `${stamp}-`));
  const stdout = fs.openSync(path.join(directory, 'stdout.log'), 'wx', 0o600);
  let stderr;
  try {
    stderr = fs.openSync(path.join(directory, 'stderr.log'), 'wx', 0o600);
  } catch (error) {
    fs.closeSync(stdout);
    throw error;
  }
  let closed = false;
  function write(kind, chunk) {
    if (closed) throw new Error('Runtime logs are already closed');
    const bytes = Buffer.isBuffer(chunk) ? chunk : Buffer.from(String(chunk));
    const fd = kind === 'stderr' ? stderr : stdout;
    for (let offset = 0; offset < bytes.length;) {
      const written = fs.writeSync(fd, bytes, offset, bytes.length - offset);
      if (!written) throw new Error('Runtime log write made no progress');
      offset += written;
    }
    mirrors[kind].write(bytes);
  }
  return {
    directory,
    stdout: chunk => write('stdout', chunk),
    stderr: chunk => write('stderr', chunk),
    info: (...args) => write('stdout', format(...args) + '\n'),
    error: (...args) => write('stderr', format(...args) + '\n'),
    close() {
      if (closed) return;
      closed = true;
      try { fs.closeSync(stdout); } finally { fs.closeSync(stderr); }
    },
  };
}

module.exports = { openRuntimeLogs };
