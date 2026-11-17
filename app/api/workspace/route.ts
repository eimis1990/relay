import { readFile, mkdir, writeFile, rename } from 'node:fs/promises';
import path from 'node:path';
import { createHmac } from 'node:crypto';
import { config, graph, json, session, setSession } from '../../../lib/outlook';
import { workspaceSchema, workspaceWriteSchema } from '../../../lib/workspace-schema';
import { databaseConfigured, readWorkspace, writeWorkspace, WorkspaceError } from '../../../lib/workspace-store';

export const runtime = 'nodejs';
export const dynamic = 'force-dynamic';
const file = path.join(process.cwd(), '.local-data', 'workspace.json');
function local(req: Request) {
  return !databaseConfigured() && process.env.NODE_ENV === 'development' && ['localhost', '127.0.0.1', '[::1]'].includes(new URL(req.url).hostname);
}
async function identity() {
  const s = await session();
  if (!s) throw new WorkspaceError('Sign in with Microsoft to open your saved workspace.', 401);
  // Upgrade older cookies; mutable email aliases never determine ownership.
  if (!s.subject) {
    const response = await graph(s.token, 'me?$select=id');
    if (!response.ok) throw new WorkspaceError('Please reconnect Microsoft to open your workspace.', 401);
    const profile = await response.json();
    if (typeof profile.id !== 'string' || !profile.id || profile.id.length > 200) throw new WorkspaceError('Could not verify your Microsoft identity. Please reconnect.', 401);
    s.subject = profile.id;
    await setSession(s);
  }
  const owner = 'microsoft:' + s.subject;
  const workspaceId = createHmac('sha256', config().secret).update('workspace:' + owner).digest('hex');
  return { owner, workspaceId, account: { email: s.email, name: s.name } };
}
function failure(error: unknown) {
  return error instanceof WorkspaceError ? json({ error: error.message }, error.status) : json({ error: 'Workspace storage is temporarily unavailable. Please retry.' }, 503);
}
export async function GET(req: Request) {
  try {
    if (local(req)) {
      let data = null;
      try { data = workspaceSchema.parse(JSON.parse(await readFile(file, 'utf8'))); }
      catch (error) { if ((error as NodeJS.ErrnoException).code !== 'ENOENT') throw error; }
      return json({ data, revision: null, workspaceId: null, mode: 'local-development', account: null });
    }
    const { owner, workspaceId, account } = await identity();
    const row = await readWorkspace(owner);
    return json({ data: row?.data ?? null, revision: row?.revision ?? null, savedAt: row?.updated_at ?? null, mode: 'cloud', workspaceId, account });
  } catch (error) { return failure(error); }
}
export async function POST(req: Request) {
  try {
    const isLocal = local(req);
    const origin = isLocal ? new URL(req.url).origin : config().origin;
    if (req.headers.get('origin') !== origin) return json({ error: 'Invalid request origin.' }, 403);
    const user = isLocal ? null : await identity();
    if (!req.headers.get('content-type')?.startsWith('application/json')) return json({ error: 'JSON content is required.' }, 415);
    const text = await req.text();
    if (Buffer.byteLength(text, 'utf8') > 24000) return json({ error: 'Request too large.' }, 413);
    let parsed;
    try { parsed = workspaceWriteSchema.safeParse(JSON.parse(text)); }
    catch { return json({ error: 'Invalid JSON.' }, 400); }
    if (!parsed.success) return json({ error: 'Check your entries. Submitted briefs require contact details and consent.' }, 400);
    const { data, revision, workspaceId } = parsed.data;
    if (!isLocal && workspaceId !== user!.workspaceId) return json({ error: 'Your signed-in account changed. Reload your workspace before saving.' }, 409);
    if (isLocal) {
      await mkdir(path.dirname(file), { recursive: true, mode: 0o700 });
      const tmp = file + '.' + crypto.randomUUID() + '.tmp';
      await writeFile(tmp, JSON.stringify(data), { mode: 0o600 });
      await rename(tmp, file);
      return json({ ok: true, data, revision: null, workspaceId: null, mode: 'local-development' });
    }
    const row = await writeWorkspace(user!.owner, data, revision);
    return json({ ok: true, data: row.data, revision: row.revision, savedAt: row.updated_at, mode: 'cloud' });
  } catch (error) { return failure(error); }
}
