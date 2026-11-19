create table public.relay_mail_accounts (
 owner_id text primary key check(length(owner_id) between 11 and 220), email text not null,
 refresh_cipher text, enabled boolean not null default false, started_at timestamptz,
 cursor_cipher text, language text not null default 'English' check(language in ('English','Lithuanian')),
 rules_cipher text, lock_id uuid, lock_until timestamptz not null default 'epoch',
 last_polled_at timestamptz, last_success_at timestamptz, last_error text,
 budget_date date not null default current_date, budget_used integer not null default 0 check(budget_used>=0)
);
create table public.relay_reviews (
 id bigint generated always as identity primary key, owner_id text not null,
 message_id text not null check(length(message_id) between 1 and 1500), snapshot text not null,
 language text not null check(language in ('English','Lithuanian')), reviewed_at timestamptz not null,
 checked boolean not null default false, priority text check(priority in ('High','Normal','Low')),
 topic text check(topic in ('Quote request','Order','Delivery','Invoice','Complaint','Supplier','General')),
 revision integer not null default 1, unique(owner_id,message_id)
);
create index relay_reviews_owner_history on public.relay_reviews(owner_id,id desc);
create table public.relay_review_queue (
 owner_id text not null references public.relay_mail_accounts(owner_id), message_id text not null,
 status text not null default 'pending' check(status in ('pending','done','failed','skipped')),
 attempts integer not null default 0, available_at timestamptz not null default now(),
 last_error text, created_at timestamptz not null default now(), primary key(owner_id,message_id)
);
create index relay_queue_pending on public.relay_review_queue(owner_id,available_at) where status='pending';
alter table public.relay_mail_accounts enable row level security;
alter table public.relay_reviews enable row level security;
alter table public.relay_review_queue enable row level security;
revoke all on public.relay_mail_accounts,public.relay_reviews,public.relay_review_queue from public,anon,authenticated,service_role;
grant select,insert,update on public.relay_mail_accounts,public.relay_reviews,public.relay_review_queue to service_role;
grant usage,select on sequence public.relay_reviews_id_seq to service_role;
comment on table public.relay_reviews is 'Private owner-scoped encrypted review snapshots, with durable human corrections. Browser roles have no access.';
comment on table public.relay_mail_accounts is 'Server-only delegated Outlook refresh tokens and delta cursors, encrypted with the Relay application key. Enable requires explicit offline consent.';

create extension if not exists pg_cron with schema pg_catalog;
create extension if not exists pg_net with schema extensions;
create schema if not exists relay_private;
revoke all on schema relay_private from public,anon,authenticated;
grant usage on schema relay_private to service_role;
-- Narrow fixed-destination scheduler installation; privileged implementation stays outside the exposed schema.
create function relay_private.configure_scheduler(p_token text) returns void
language plpgsql security definer set search_path='' as $$
declare sid uuid;
begin
 if length(p_token)<>64 or p_token !~ '^[a-f0-9]+$' then raise exception 'Invalid scheduler token'; end if;
 select id into sid from vault.secrets where name='relay_review_scheduler';
 if sid is null then perform vault.create_secret(p_token,'relay_review_scheduler');
 else perform vault.update_secret(sid,p_token); end if;
 perform cron.schedule('relay-inbox-review','*/5 * * * *',$job$
 select net.http_post(
 url:='https://relay-navy-alpha.vercel.app/api/cron/review',
 headers:=jsonb_build_object('Content-Type','application/json','Authorization','Bearer '||(select decrypted_secret from vault.decrypted_secrets where name='relay_review_scheduler')),
 body:='{}'::jsonb,timeout_milliseconds:=180000
 ) where exists(select 1 from public.relay_mail_accounts where enabled);
 $job$);
end $$;
revoke all on function relay_private.configure_scheduler(text) from public,anon,authenticated;
grant execute on function relay_private.configure_scheduler(text) to service_role;
create function public.relay_configure_scheduler(p_token text) returns void
language sql security invoker set search_path='' as $$ select relay_private.configure_scheduler(p_token); $$;
revoke all on function public.relay_configure_scheduler(text) from public,anon,authenticated;
grant execute on function public.relay_configure_scheduler(text) to service_role;
