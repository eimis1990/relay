import 'server-only';
import { z } from 'zod';
import { workspaceSchema, type Workspace } from './workspace-schema';

export class WorkspaceError extends Error {
  constructor(message: string, public status: number) { super(message); }
}
export function databaseConfigured() {
  return Boolean(process.env.SUPABASE_URL && process.env.SUPABASE_SECRET_KEY);
}
const rowSchema = z.object({
  data: workspaceSchema, revision: z.number().int().positive(), updated_at: z.string(),
});
async function request(query: URLSearchParams, init: RequestInit = {}) {
  if (!databaseConfigured()) throw new WorkspaceError('Workspace storage is not configured yet. Please contact the Relay owner.', 503);
  const base = new URL(process.env.SUPABASE_URL!);
  // Loopback supports Supabase CLI development and integration tests only.
  const local = base.protocol === 'http:' && ['127.0.0.1', 'localhost'].includes(base.hostname);
  if ((!local && base.protocol !== 'https:') || base.username || base.password || base.pathname !== '/') {
    throw new WorkspaceError('Workspace storage configuration is invalid.', 503);
  }
  const key = process.env.SUPABASE_SECRET_KEY!;
  const response = await fetch(new URL('/rest/v1/relay_workspaces?' + query, base), {
    ...init,
    headers: {
      apikey: key,
      // Legacy service_role JWTs need Authorization; new sb_secret keys do not.
      ...(!key.startsWith('sb_secret_') ? { Authorization: 'Bearer ' + key } : {}),
      'Content-Type': 'application/json', Prefer: 'return=representation',
    },
    cache: 'no-store', signal: AbortSignal.timeout(10000),
  });
  if (response.status === 409) throw new WorkspaceError('This workspace changed in another tab. Reload before saving again.', 409);
  if (!response.ok) throw new WorkspaceError('Workspace storage is temporarily unavailable. Your changes have not been saved; please retry.', 503);
  return z.array(rowSchema).parse(await response.json());
}
export async function readWorkspace(owner: string) {
  const rows = await request(new URLSearchParams({ owner_id: 'eq.' + owner, select: 'data,revision,updated_at', limit: '1' }));
  return rows[0] ?? null;
}
export async function writeWorkspace(owner: string, data: Workspace, revision: number | null) {
  const query = new URLSearchParams({ select: 'data,revision,updated_at' });
  if (revision !== null) {
    query.set('owner_id', 'eq.' + owner);
    query.set('revision', 'eq.' + revision);
  }
  // Insert never upserts. Update compares the revision in the same database statement.
  const rows = await request(query, {
    method: revision === null ? 'POST' : 'PATCH',
    body: JSON.stringify({ ...(revision === null ? { owner_id: owner } : {}), data, revision: (revision ?? 0) + 1, updated_at: new Date().toISOString() }),
  });
  if (rows.length !== 1) throw new WorkspaceError('This workspace changed in another tab. Reload before saving again.', 409);
  return rows[0];
}
