create table public.relay_workflows (
 owner_id text primary key check(length(owner_id) between 11 and 220),
 settings_cipher text not null check(octet_length(settings_cipher)<40000),
 revision integer not null default 1 check(revision>0), updated_at timestamptz not null default now()
);
alter table public.relay_workflows enable row level security;
revoke all on public.relay_workflows from public,anon,authenticated,service_role;
grant select,insert,update on public.relay_workflows to service_role;
alter table public.relay_review_queue add column received_at timestamptz not null default now();
update public.relay_review_queue set received_at=created_at;
comment on table public.relay_workflows is 'Encrypted per-Microsoft-subject reviewer configuration. Only server routes can access; changes require a matching revision.';
