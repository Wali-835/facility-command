-- Facility Command — Phase 6: parts catalog responsibility + filtered reporting
-- Also includes Phase 7's PO-for-repair (the one item not needing external scoping).
-- Run this once in the Supabase SQL Editor. Safe to re-run (IF NOT EXISTS).

-- ============================================================
-- 1. Responsible user per model (catalog) and per category.
-- ============================================================
alter table mhe_models add column if not exists responsible_name text;
alter table mhe_models add column if not exists responsible_email text;

create table if not exists category_responsibility (
  id text primary key,
  category text not null unique,
  responsible_name text,
  responsible_email text
);
alter table category_responsibility enable row level security;
drop policy if exists "category_responsibility_all_authenticated" on category_responsibility;
create policy "category_responsibility_all_authenticated" on category_responsibility for all to authenticated using (true) with check (true);

-- ============================================================
-- 2. PO-for-repair on work orders (mirrors the PO fields already on assets).
-- ============================================================
alter table work_orders add column if not exists po_number text;
alter table work_orders add column if not exists po_amount numeric;
alter table work_orders add column if not exists po_status text; -- 'Requested' | 'Issued' | 'Received'

notify pgrst, 'reload schema';
