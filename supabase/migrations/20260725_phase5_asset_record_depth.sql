-- Facility Command — Phase 5: Asset record depth
-- Run this once in the Supabase SQL Editor. Safe to re-run (IF NOT EXISTS).

-- ============================================================
-- 1. Category-specific fields — flexible JSONB instead of bolting on
--    columns per category (HVAC set point, electrical voltage, etc.)
-- ============================================================
alter table assets add column if not exists custom_fields jsonb not null default '{}'::jsonb;

-- ============================================================
-- 2. Document storage (manuals, IQ/OQ certificates, etc.)
--    Same pattern as the work-order-photos bucket, plus a metadata table
--    so documents can carry a type/description (unlike plain photos).
-- ============================================================
create table if not exists asset_documents (
  id text primary key,
  asset_id text not null,
  asset_name text,
  document_type text, -- 'Manual' | 'IQ' | 'OQ' | 'Certificate' | 'Other'
  file_name text not null,
  file_path text not null, -- path inside the asset-documents storage bucket
  notes text,
  uploaded_by text,
  uploaded_at timestamptz not null default now()
);
create index if not exists idx_asset_documents_asset_id on asset_documents(asset_id);
alter table asset_documents enable row level security;
drop policy if exists "asset_documents_all_authenticated" on asset_documents;
create policy "asset_documents_all_authenticated" on asset_documents for all to authenticated using (true) with check (true);

insert into storage.buckets (id, name, public)
values ('asset-documents', 'asset-documents', true)
on conflict (id) do nothing;

drop policy if exists "asset_documents_storage_read" on storage.objects;
create policy "asset_documents_storage_read" on storage.objects for select to public using (bucket_id = 'asset-documents');
drop policy if exists "asset_documents_storage_write" on storage.objects;
create policy "asset_documents_storage_write" on storage.objects for insert to authenticated with check (bucket_id = 'asset-documents');
drop policy if exists "asset_documents_storage_delete" on storage.objects;
create policy "asset_documents_storage_delete" on storage.objects for delete to authenticated using (bucket_id = 'asset-documents');

-- ============================================================
-- 3. Insurance policies + expiry tracking
-- ============================================================
create table if not exists insurance_policies (
  id text primary key,
  asset_id text not null,
  asset_name text,
  provider text,
  policy_number text,
  coverage_type text,
  start_date date,
  expiry_date date,
  premium numeric,
  notes text,
  created_at timestamptz not null default now()
);
create index if not exists idx_insurance_policies_asset_id on insurance_policies(asset_id);
create index if not exists idx_insurance_policies_expiry on insurance_policies(expiry_date);
alter table insurance_policies enable row level security;
drop policy if exists "insurance_policies_all_authenticated" on insurance_policies;
create policy "insurance_policies_all_authenticated" on insurance_policies for all to authenticated using (true) with check (true);

notify pgrst, 'reload schema';
