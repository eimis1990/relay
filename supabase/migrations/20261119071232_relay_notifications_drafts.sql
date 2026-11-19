create table public.relay_telegram (
 owner_id text primary key,
 id uuid not null unique default gen_random_uuid(),
 bot_id text unique,
 bot_username text,
 token_cipher text,
 webhook_hash text,
 link_hash text,
 link_expires timestamptz,
 pending_chat_cipher text,
 pending_name text,
 chat_cipher text,
 chat_name text,
 enabled boolean not null default false,
 preferences jsonb not null default '{"allReviews":true,"highPriority":false,"needsAttention":false,"draftReady":false,"includeSummary":false}',
 last_error text,
 last_sent_at timestamptz,
 updated_at timestamptz not null default now()
);
alter table public.relay_telegram enable row level security;
revoke all on public.relay_telegram from public, anon, authenticated, service_role;
grant select,insert,update on public.relay_telegram to service_role;
alter table public.relay_reviews
 add column draft_status text not null default 'none' check(draft_status in ('none','pending','generating','ready','failed')),
 add column draft_cipher text,
 add column draft_attempts integer not null default 0,
 add column draft_error text,
 add column draft_revision integer not null default 0,
 add column draft_available_at timestamptz not null default now(),
 add column draft_outlook_status text not null default 'none' check(draft_outlook_status in ('none','pending','creating','done','failed','uncertain')),
 add column draft_outlook_cipher text,
 add column notify_status text not null default 'none' check(notify_status in ('none','pending','sending','done','failed','uncertain','skipped')),
 add column notify_attempts integer not null default 0,
 add column notify_available_at timestamptz not null default now(),
 add column notify_error text;
create index relay_drafts_pending on public.relay_reviews(owner_id,draft_available_at) where draft_status in ('pending','generating');
create index relay_notify_pending on public.relay_reviews(owner_id,notify_available_at) where notify_status in ('pending','sending');
