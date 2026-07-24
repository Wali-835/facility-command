-- Facility Command — Purchase Orders (procurement), replacing the earlier
-- one-PO-per-work-order fields with a proper many-parts-per-PO model,
-- same shape as insurance_policies / insurance_policy_assets: one PO can
-- cover several line items across different models and assets.
-- Run this once in the Supabase SQL Editor. Safe to re-run (IF NOT EXISTS).
--
-- The work_orders.po_number/po_amount/po_status columns added in
-- 20260729_phase6_parts_responsibility_reporting.sql are no longer used by
-- the app and are left in place (harmless) rather than dropped.

create table if not exists purchase_orders (
  id text primary key,
  po_number text,
  vendor text not null,
  status text default 'Requested', -- 'Requested' | 'Issued' | 'Received'
  order_date date,
  notes text,
  created_at timestamptz default now()
);
alter table purchase_orders enable row level security;
drop policy if exists "purchase_orders_all_authenticated" on purchase_orders;
create policy "purchase_orders_all_authenticated" on purchase_orders for all to authenticated using (true) with check (true);

create table if not exists purchase_order_items (
  id text primary key,
  po_id text not null,
  spare_part_id text, -- links back to spare_parts when this line item came from an approved "Order" request
  part_name text not null,
  part_number text,
  model text,
  asset_id text,
  asset_name text,
  quantity numeric default 1,
  unit_cost numeric,
  total_cost numeric
);
create index if not exists idx_poi_po on purchase_order_items(po_id);
create index if not exists idx_poi_spare_part on purchase_order_items(spare_part_id);
alter table purchase_order_items enable row level security;
drop policy if exists "purchase_order_items_all_authenticated" on purchase_order_items;
create policy "purchase_order_items_all_authenticated" on purchase_order_items for all to authenticated using (true) with check (true);

notify pgrst, 'reload schema';
