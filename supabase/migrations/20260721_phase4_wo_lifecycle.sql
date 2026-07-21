-- Facility Command — Phase 4: WO/breakdown linkage, technician assignment,
-- spare-parts approval (new Engineer role), ticketing system.
-- Run this once in the Supabase SQL Editor. Safe to re-run (IF NOT EXISTS).

-- ============================================================
-- 1. Real WO <-> breakdown/issue linkage (replaces matching WOs by
--    asset name + status, which could grab the wrong work order).
-- ============================================================
alter table work_orders add column if not exists breakdown_id text;
alter table work_orders add column if not exists issue_id text;
create index if not exists idx_work_orders_breakdown_id on work_orders(breakdown_id);
create index if not exists idx_work_orders_issue_id on work_orders(issue_id);

-- Mirrors issue_reports.work_order_id, which already exists and works well.
alter table breakdown_reports add column if not exists work_order_id text;

-- ============================================================
-- 2. Spare-parts approval on work orders.
--    Existing parts are treated as already-approved (they were already
--    deducted from stock under the old implicit-approval behavior).
--    New parts default to Pending and are set explicitly by the app.
-- ============================================================
alter table spare_parts add column if not exists approval_status text not null default 'Approved';
alter table spare_parts add column if not exists approved_by text;
alter table spare_parts add column if not exists approved_at timestamptz;
alter table spare_parts add column if not exists rejection_reason text;

-- ============================================================
-- 3. Ticketing system — general requests, optionally linked to a
--    work order and/or an asset, with its own event/comment log.
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
-- 4. Engineer role — approves/rejects spare-parts requests.
--    user_roles.role has no CHECK constraint in this schema, so no
--    migration is needed to "allow" the new value. If your database
--    does have one, this statement finds and drops it; harmless if not.
-- ============================================================
do $$
declare
  con text;
begin
  select conname into con
  from pg_constraint
  where conrelid = 'user_roles'::regclass and contype = 'c'
  and pg_get_constraintdef(oid) ilike '%role%';
  if con is not null then
    execute format('alter table user_roles drop constraint %I', con);
  end if;
end $$;
