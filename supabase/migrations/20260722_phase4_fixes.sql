-- Facility Command — Phase 4 fixes
-- Run this once in the Supabase SQL Editor. Safe to re-run (IF NOT EXISTS).
--
-- If you ran 20260721_phase4_wo_lifecycle.sql and the Tickets tab says
-- "Could not find the table 'public.tickets'", that migration's very last
-- step (a defensive block that tried to drop a role CHECK constraint that
-- probably never existed) most likely errored out and silently rolled back
-- EVERYTHING in that script, including the tickets tables. This migration
-- re-does everything from that file (harmless if it already applied) MINUS
-- that risky last step, and explicitly tells PostgREST to reload its schema
-- cache so new tables are picked up immediately instead of on a delay.

-- ============================================================
-- 1. Real WO <-> breakdown/issue linkage
-- ============================================================
alter table work_orders add column if not exists breakdown_id text;
alter table work_orders add column if not exists issue_id text;
create index if not exists idx_work_orders_breakdown_id on work_orders(breakdown_id);
create index if not exists idx_work_orders_issue_id on work_orders(issue_id);
alter table breakdown_reports add column if not exists work_order_id text;

-- ============================================================
-- 2. Spare-parts approval — now a two-stage workflow:
--    maintenance requests -> supervisor approves -> engineer gives
--    final approval (and confirms/sets the cost).
-- ============================================================
alter table spare_parts add column if not exists approval_status text not null default 'Approved';
alter table spare_parts add column if not exists supervisor_approved_by text;
alter table spare_parts add column if not exists supervisor_approved_at timestamptz;
alter table spare_parts add column if not exists approved_by text;
alter table spare_parts add column if not exists approved_at timestamptz;
alter table spare_parts add column if not exists rejection_reason text;

-- ============================================================
-- 3. Ticketing system
-- ============================================================
create table if not exists tickets (
  id text primary key,
  title text not null,
  description text,
  category text,
  priority text not null default 'Medium',
  status text not null default 'Open',
  site text,
  asset_id text,
  asset_name text,
  work_order_id text,
  requested_by text,
  requested_at timestamptz not null default now(),
  assignee text
);
alter table tickets enable row level security;
drop policy if exists "tickets_all_authenticated" on tickets;
create policy "tickets_all_authenticated" on tickets for all to authenticated using (true) with check (true);

create table if not exists ticket_events (
  id text primary key,
  ticket_id text not null,
  event_type text not null, -- 'created' | 'comment' | 'status_change'
  note text,
  by text,
  at timestamptz not null default now()
);
create index if not exists idx_ticket_events_ticket_id on ticket_events(ticket_id);
alter table ticket_events enable row level security;
drop policy if exists "ticket_events_all_authenticated" on ticket_events;
create policy "ticket_events_all_authenticated" on ticket_events for all to authenticated using (true) with check (true);

-- ============================================================
-- 4. In-app notifications (e.g. "you were assigned WO-XXXX")
-- ============================================================
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

-- ============================================================
-- 5. Make sure PostgREST knows about all of the above right away.
-- ============================================================
notify pgrst, 'reload schema';
