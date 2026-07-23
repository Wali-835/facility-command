-- Facility Command — Asset code + subcategory
-- Run this once in the Supabase SQL Editor. Safe to re-run (IF NOT EXISTS).

alter table assets add column if not exists subcategory text;
alter table assets add column if not exists asset_code text;

-- Internal asset code should be unique per asset, same pattern as serial_number.
create unique index if not exists assets_asset_code_unique
  on assets (asset_code)
  where asset_code is not null and asset_code <> '';

notify pgrst, 'reload schema';
