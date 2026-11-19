alter table public.relay_mail_accounts
  add column category_enabled boolean not null default false,
  add column category_permission boolean not null default false,
  add column category_error text,
  add column category_last_success timestamptz;
-- Durable outbox lives on the review, so saving AI results and scheduling its
-- category update are atomic. Existing reviews are deliberately not backfilled.
alter table public.relay_reviews
  add column category_status text not null default 'none' check (category_status in ('none','pending','done','failed','skipped')),
  add column category_payload text,
  add column category_revision uuid not null default gen_random_uuid(),
  add column category_attempts integer not null default 0 check (category_attempts >= 0),
  add column category_available_at timestamptz not null default now(),
  add column category_error text,
  add column category_synced_at timestamptz;
create index relay_categories_pending on public.relay_reviews(owner_id, category_available_at) where category_status='pending';
-- Existing owner-isolated server access only; no browser role can read tokens,
-- email snapshots, or schedule mailbox writes directly.
alter table public.relay_mail_accounts enable row level security;
alter table public.relay_reviews enable row level security;
revoke all on public.relay_mail_accounts, public.relay_reviews from public, anon, authenticated;
grant select, insert, update on public.relay_mail_accounts, public.relay_reviews to service_role;
