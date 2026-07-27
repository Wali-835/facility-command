-- Facility Command — Ensure the notifications table exists
-- The browser console showed a 404 fetching /rest/v1/notifications, which
-- means PostgREST doesn't see this table — either it was never created on
-- this database, or it exists but PostgREST's schema cache was never
-- refreshed after it was added. This just re-asserts it (safe no-op if it
-- already exists) and forces a schema reload either way.
-- Run this once in the Supabase SQL Editor. Safe to re-run.

create table if not exists notifications (
  id text primary key,
  recipient text not null, -- matches user_roles.name
  type text not null,
  message text not null,
  link_type text, -- 'work_order' | 'ticket' | etc.
  link_id text,
  read boolean not null default false,
  created_at timestamptz not null default now()
);
create index if not exists idx_notifications_recipient on notifications(recipient);
alter table notifications enable row level security;
drop policy if exists "notifications_all_authenticated" on notifications;
create policy "notifications_all_authenticated" on notifications for all to authenticated using (true) with check (true);

notify pgrst, 'reload schema';
