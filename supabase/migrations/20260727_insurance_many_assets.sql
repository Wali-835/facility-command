-- Facility Command — Insurance: one policy can cover many assets
-- Run this once in the Supabase SQL Editor. Safe to re-run (IF NOT EXISTS).
--
-- insurance_policies used to be one-row-per-asset. This adds a join table
-- so a single policy can cover any number of assets, plus a document
-- attachment on the policy itself (stored in the existing asset-documents
-- bucket under a policy-docs/ prefix — no new bucket needed).
-- Existing per-asset policy rows are migrated into the join table so
-- nothing already entered is lost; the old asset_id/asset_name columns on
-- insurance_policies are left in place (unused by the app going forward)
-- rather than dropped, to avoid any data loss risk.

alter table insurance_policies add column if not exists file_name text;
alter table insurance_policies add column if not exists file_path text;

create table if not exists insurance_policy_assets (
  id text primary key,
  policy_id text not null,
  asset_id text not null,
  asset_name text
);
create index if not exists idx_ipa_policy on insurance_policy_assets(policy_id);
create index if not exists idx_ipa_asset on insurance_policy_assets(asset_id);
alter table insurance_policy_assets enable row level security;
drop policy if exists "insurance_policy_assets_all_authenticated" on insurance_policy_assets;
create policy "insurance_policy_assets_all_authenticated" on insurance_policy_assets for all to authenticated using (true) with check (true);

-- Backfill: any existing policy that already had a single asset_id gets a
-- matching row in the new join table (skipped if already backfilled).
insert into insurance_policy_assets (id, policy_id, asset_id, asset_name)
select 'IPA-BACKFILL-' || p.id, p.id, p.asset_id, p.asset_name
from insurance_policies p
where p.asset_id is not null
  and not exists (
    select 1 from insurance_policy_assets ipa
    where ipa.policy_id = p.id and ipa.asset_id = p.asset_id
  );

notify pgrst, 'reload schema';
