/**
 * Host half intentionally has no filesystem or session mutation capability.
 * The authenticated onboarding route is owned by the deployment host plugin.
 */
import { checkedDirectory, ensureLegacySessions, readSetup, sameOrChild, writeSetup } from './state.js';
export const inject = ['connection', 'sessionController'];

async function inspectDirectory(ctx, sessionId) {
  const inspection = await ctx.sessionController.inspect(sessionId);
  const cwd = inspection?.meta?.cwd;
  if (typeof cwd !== 'string' || !cwd) throw new Error('会话没有已确定的目录');
  return checkedDirectory(cwd);
}

function json(value, status = 200) {
  return Response.json(value, { status, headers: { 'cache-control': 'no-store' } });
}

export async function handleOnboarding(ctx, request) {
  try {
    if (request.method === 'GET') {
      const sessionId = new URL(request.url).searchParams.get('session_id');
      if (!sessionId) return json({ error: '缺少会话编号' }, 400);
      const sessionDirectory = await inspectDirectory(ctx, sessionId);
      const record = await readSetup(sessionId);
      if (!record) return json({ status: 'pending' });
      const officeWorkspace = await checkedDirectory(record.officeWorkspace);
      if (record.sessionDirectory !== sessionDirectory || officeWorkspace !== record.officeWorkspace || !sameOrChild(sessionDirectory, officeWorkspace)) return json({ error: '确认记录与会话目录不一致' }, 409);
      return json(record);
    }
    const body = await request.json();
    const sessionId = body?.session_id;
    if (typeof sessionId !== 'string') return json({ error: '缺少会话编号' }, 400);
    const sessionDirectory = await inspectDirectory(ctx, sessionId);
    const chosenSessionDirectory = await checkedDirectory(body?.session_directory);
    if (chosenSessionDirectory !== sessionDirectory) return json({ error: '所选会话目录与此会话不符' }, 409);
    const officeWorkspace = await checkedDirectory(body?.office_workspace);
    if (!sameOrChild(sessionDirectory, officeWorkspace)) return json({ error: '办公工作区必须等于会话目录或位于其子目录' }, 400);
    const record = { version: 1, sessionId, sessionDirectory, officeWorkspace, confirmedAt: new Date().toISOString() };
    await writeSetup(record);
    return json(record);
  } catch (error) {
    return json({ error: `会话目录确认失败：${error instanceof Error ? error.message : String(error)}` }, 503);
  }
}

export async function apply(ctx) {
  await ensureLegacySessions();
  ctx.connection.fetch.register({
    path: '/api/onboarding',
    methods: ['GET', 'POST'],
    requestBody: 'buffered',
    fetch: (request) => handleOnboarding(ctx, request)
  });
}
