-- Facility Command — Stock adjustment requests (add/deduct model-part stock)
-- Supervisor/engineer requests a manual add or deduct against a model part's
-- stock_quantity; an engineer or admin approves, which then applies the
-- change. Plain stock edits by an admin still bypass this (direct edit
-- stays available to admins in the Parts Catalog Mgmt screen).
-- Run this once in the Supabase SQL Editor. Safe to re-run (IF NOT EXISTS).

create table if not exists stock_adjustments (
  id text primary key,
  part_id text not null, -- model_parts.id
  part_name text,
  model text,
  adjustment_type text not null, -- 'Add' | 'Deduct'
  quantity numeric not null,
  reason text,
  requested_by text,
  requested_at timestamptz default now(),
  status text default 'Pending', -- 'Pending' | 'Approved' | 'Rejected'
  approved_by text,
  approved_at timestamptz
);
create index if not exists idx_stock_adj_part on stock_adjustments(part_id);
create index if not exists idx_stock_adj_status on stock_adjustments(status);
alter table stock_adjustments enable row level security;
drop policy if exists "stock_adjustments_all_authenticated" on stock_adjustments;
create policy "stock_adjustments_all_authenticated" on stock_adjustments for all to authenticated using (true) with check (true);

notify pgrst, 'reload schema';
