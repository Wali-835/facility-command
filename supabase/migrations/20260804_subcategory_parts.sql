-- Facility Command — Category/Subcategory spare parts catalog
-- The parts catalog was keyed to mhe_models, which only fits equipment
-- that genuinely has brand/model spare parts (MHE). Other categories
-- (HVAC, Electrical, Plumbing, etc.) need their own independent parts
-- lists scoped to just Category + Subcategory, with no equipment model
-- required. This runs alongside the existing model_parts catalog rather
-- than replacing it.
-- Run this once in the Supabase SQL Editor. Safe to re-run (IF NOT EXISTS).

create table if not exists subcategory_parts (
  id text primary key,
  category text not null,
  subcategory text not null,
  part_name text not null,
  part_number text,
  supplier text,
  unit_cost numeric default 0,
  stock_quantity numeric default 0,
  min_stock_level numeric default 1,
  notes text
);
create index if not exists idx_subcat_parts_scope on subcategory_parts(category, subcategory);
alter table subcategory_parts enable row level security;
drop policy if exists "subcategory_parts_all_authenticated" on subcategory_parts;
create policy "subcategory_parts_all_authenticated" on subcategory_parts for all to authenticated using (true) with check (true);

-- Stock adjustment requests need to know whether they're adjusting a
-- model_parts row or a subcategory_parts row.
alter table stock_adjustments add column if not exists part_source text default 'model'; -- 'model' | 'subcategory'

notify pgrst, 'reload schema';
