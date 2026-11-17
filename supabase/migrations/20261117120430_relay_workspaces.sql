-- One private pilot workspace per verified Microsoft Graph identity.
-- Access only through Relay's server routes; never expose the service key to clients.
create table public.relay_workspaces (
  owner_id text primary key check (length(owner_id) between 11 and 220),
  data jsonb not null check (jsonb_typeof(data) = 'object' and octet_length(data::text) <= 32768),
  revision integer not null default 1 check (revision > 0),
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);
alter table public.relay_workspaces enable row level security;
revoke all on table public.relay_workspaces from public, anon, authenticated, service_role;
grant select, insert, update on table public.relay_workspaces to service_role;
comment on table public.relay_workspaces is 'Relay owner-scoped onboarding briefs. No OAuth credentials or email contents. RLS denies browser roles; backend verifies Microsoft session and always filters by owner_id.';
