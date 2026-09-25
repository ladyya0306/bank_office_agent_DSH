import { mkdir, readFile, readdir, realpath, rename, stat, unlink, writeFile } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const workspaceHome = process.env.DSH_HOME || path.resolve(path.dirname(fileURLToPath(import.meta.url)), '../../..');
export const stateDirectory = path.join(workspaceHome, 'session-office-setup');
export const legacyFile = path.join(stateDirectory, 'legacy-sessions.json');
const sessionIdPattern = /^[A-Za-z0-9_-]{8,128}$/;

export function stateFile(sessionId) {
  if (typeof sessionId !== 'string' || !sessionIdPattern.test(sessionId)) throw new Error('会话编号无效');
  return path.join(stateDirectory, `${sessionId}.json`);
}

export async function checkedDirectory(value) {
  if (typeof value !== 'string' || !path.isAbsolute(value)) throw new Error('目录必须是绝对路径');
  const resolved = await realpath(value);
  if (!(await stat(resolved)).isDirectory()) throw new Error('所选路径不是目录');
  return resolved;
}

export function sameOrChild(parent, candidate) {
  const relative = path.relative(parent, candidate);
  return relative === '' || (relative !== '..' && !relative.startsWith(`..${path.sep}`) && !path.isAbsolute(relative));
}

export async function readSetup(sessionId) {
  const file = stateFile(sessionId);
  try {
    const record = JSON.parse(await readFile(file, 'utf8'));
    if (record.version !== 1 || record.sessionId !== sessionId || typeof record.sessionDirectory !== 'string' || typeof record.officeWorkspace !== 'string') throw new Error('确认记录格式无效');
    return record;
  } catch (error) {
    if (error?.code === 'ENOENT') return null;
    throw error;
  }
}

export async function writeSetup(record) {
  const file = stateFile(record.sessionId);
  await mkdir(stateDirectory, { recursive: true });
  const temporary = `${file}.${process.pid}.${crypto.randomUUID()}.tmp`;
  try {
    await writeFile(temporary, JSON.stringify(record), { encoding: 'utf8', flag: 'wx' });
    await rename(temporary, file);
  } catch (error) {
    await unlink(temporary).catch(() => {});
    throw error;
  }
}

export async function ensureLegacySessions() {
  await mkdir(stateDirectory, { recursive: true });
  try {
    await stat(legacyFile);
    return;
  } catch (error) {
    if (error?.code !== 'ENOENT') throw error;
  }
  const ids = new Set();
  const cacheDirectory = path.join(workspaceHome, 'storages', 'session_projcache', 'sessions');
  for (const entry of await readdir(cacheDirectory).catch((error) => {
    if (error?.code === 'ENOENT') return [];
    throw error;
  })) {
    if (entry.endsWith('.json')) {
      const id = entry.slice(0, -5);
      if (!sessionIdPattern.test(id)) continue;
      const cached = JSON.parse(await readFile(path.join(cacheDirectory, entry), 'utf8'));
      if (cached?.record?.rows?.sessionListMetadata?.val?.blank === false) ids.add(id);
    }
  }
  try {
    await writeFile(legacyFile, JSON.stringify({ version: 1, sessionIds: [...ids].sort() }), { encoding: 'utf8', flag: 'wx' });
  } catch (error) {
    if (error?.code !== 'EEXIST') throw error;
  }
}
